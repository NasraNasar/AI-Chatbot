import os
from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st
from docx import Document
from dotenv import load_dotenv
from google import genai
from google.genai import types
from openpyxl import load_workbook
from pptx import Presentation
from pypdf import PdfReader

from auth import validate_email, validate_name, validate_password
from database import (
    authenticate_user,
    clear_conversation,
    create_conversation,
    create_tables,
    create_user,
    delete_conversation,
    get_conversation,
    get_conversations,
    get_messages,
    remove_document,
    save_document,
    save_message,
    update_conversation_title,
)


st.set_page_config(
    page_title="Nova AI",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

load_dotenv()
create_tables()


def load_css():
    css_path = Path(__file__).with_name("styles.css")
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)


load_css()


# ==================================================
# AUTHENTICATION
# ==================================================

st.session_state.setdefault("logged_in", False)
st.session_state.setdefault("user", None)


def logout_user():
    for key in ("conversation_id", "messages", "document_name", "document_text"):
        st.session_state.pop(key, None)
    st.session_state.logged_in = False
    st.session_state.user = None


def display_authentication_page():
    st.markdown(
        '<div class="auth-container"><div class="auth-logo">✣</div>'
        '<div class="auth-title">Welcome to Nova AI</div>'
        '<div class="auth-subtitle">Sign in to continue to your personal AI assistant.</div></div>',
        unsafe_allow_html=True,
    )

    _, auth_column, _ = st.columns([1, 1.25, 1])
    with auth_column:
        login_tab, register_tab = st.tabs(["Sign in", "Create account"])

        with login_tab:
            with st.form("login_form"):
                email = st.text_input("Email address", placeholder="name@example.com")
                password = st.text_input("Password", type="password", placeholder="Enter your password")
                submitted = st.form_submit_button("Sign in", use_container_width=True, type="primary")

            if submitted:
                if not email or not password:
                    st.error("Enter your email and password.")
                else:
                    user = authenticate_user(email, password)
                    if user:
                        st.session_state.logged_in = True
                        st.session_state.user = user
                        st.rerun()
                    else:
                        st.error("Incorrect email or password.")

        with register_tab:
            with st.form("registration_form"):
                name = st.text_input("Full name", placeholder="Enter your name")
                new_email = st.text_input("Email address", placeholder="name@example.com", key="register_email")
                new_password = st.text_input("Password", type="password", placeholder="At least 8 characters", key="register_password")
                confirm_password = st.text_input("Confirm password", type="password")
                register = st.form_submit_button("Create account", use_container_width=True, type="primary")

            if register:
                name_ok, name_error = validate_name(name)
                email_ok, email_error = validate_email(new_email)
                password_ok, password_error = validate_password(new_password)

                if not name_ok:
                    st.error(name_error)
                elif not email_ok:
                    st.error(email_error)
                elif not password_ok:
                    st.error(password_error)
                elif new_password != confirm_password:
                    st.error("Passwords do not match.")
                else:
                    result = create_user(name, new_email, new_password)
                    if result["success"]:
                        st.session_state.logged_in = True
                        st.session_state.user = authenticate_user(new_email, new_password)
                        st.rerun()
                    else:
                        st.error(result["message"])


if not st.session_state.logged_in:
    display_authentication_page()
    st.stop()

current_user_id = st.session_state.user["id"]


# ==================================================
# GEMINI AND FILE PROCESSING
# ==================================================

# First, try reading values from the local .env file
api_key = os.getenv("GEMINI_API_KEY")
model_name = os.getenv("GEMINI_MODEL")

# If not available locally, try Streamlit Cloud Secrets
if not api_key:
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except (KeyError, FileNotFoundError):
        api_key = None

if not model_name:
    try:
        model_name = st.secrets.get(
            "GEMINI_MODEL",
            "gemini-3.1-flash-lite",
        )
    except FileNotFoundError:
        model_name = "gemini-3.1-flash-lite"

# Stop the application if no API key was provided
if not api_key:
    st.error(
        "Gemini API key is missing. Add it to your .env file "
        "or Streamlit Cloud Secrets."
    )
    st.stop()

# Create the Gemini client
client = genai.Client(api_key=api_key)


def extract_pdf_text(data):
    reader = PdfReader(BytesIO(data))
    return "\n\n".join(
        f"--- PDF Page {number} ---\n{text}"
        for number, page in enumerate(reader.pages, 1)
        if (text := page.extract_text())
    ).strip()


def extract_word_text(data):
    document = Document(BytesIO(data))
    content = [p.text.strip() for p in document.paragraphs if p.text.strip()]
    for number, table in enumerate(document.tables, 1):
        content.append(f"--- Word Table {number} ---")
        content.extend(" | ".join(cell.text.strip() for cell in row.cells) for row in table.rows)
    return "\n".join(content).strip()


def extract_excel_text(data):
    workbook = load_workbook(BytesIO(data), data_only=True, read_only=True)
    content = []
    for sheet in workbook.worksheets:
        content.append(f"--- Excel Sheet: {sheet.title} ---")
        for row in sheet.iter_rows(values_only=True):
            values = ["" if value is None else str(value) for value in row]
            if any(value.strip() for value in values):
                content.append(" | ".join(values))
    workbook.close()
    return "\n".join(content).strip()


def extract_powerpoint_text(data):
    presentation = Presentation(BytesIO(data))
    content = []
    for number, slide in enumerate(presentation.slides, 1):
        content.append(f"--- PowerPoint Slide {number} ---")
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                content.append(shape.text.strip())
            if getattr(shape, "has_table", False):
                content.extend(" | ".join(cell.text.strip() for cell in row.cells) for row in shape.table.rows)
    return "\n".join(content).strip()


def analyze_image(data, mime_type):
    if not client:
        raise ValueError("Gemini API key is required to analyze images.")
    if len(data) > 15 * 1024 * 1024:
        raise ValueError("Use an image smaller than 15 MB.")
    response = client.models.generate_content(
        model=model_name,
        contents=[
            "Describe this image carefully, including visible text, objects, people, tables, charts and context.",
            types.Part.from_bytes(data=data, mime_type=mime_type),
        ],
    )
    return response.text or "No image description was generated."


def extract_uploaded_file(uploaded_file):
    extension = uploaded_file.name.rsplit(".", 1)[-1].lower()
    data = uploaded_file.getvalue()
    if extension == "pdf":
        content = extract_pdf_text(data)
    elif extension == "docx":
        content = extract_word_text(data)
    elif extension == "xlsx":
        content = extract_excel_text(data)
    elif extension == "pptx":
        content = extract_powerpoint_text(data)
    elif extension in {"png", "jpg", "jpeg", "webp"}:
        mime_type = uploaded_file.type or ("image/jpeg" if extension in {"jpg", "jpeg"} else f"image/{extension}")
        content = analyze_image(data, mime_type)
    else:
        raise ValueError(f"Unsupported file type: .{extension}")
    if not content.strip():
        raise ValueError("No readable content was found in this file.")
    return content


# ==================================================
# CHAT HELPERS AND SESSION
# ==================================================

def escape_html(value):
    return str(value or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def make_title(message):
    cleaned = " ".join(message.split())
    return cleaned if len(cleaned) <= 30 else cleaned[:30] + "..."


def conversation_matches_search(conversation, search_term):
    normalized_term = search_term.strip().casefold()
    if not normalized_term:
        return True

    searchable_text = [
        conversation["title"],
        conversation["document_name"] or "",
    ]
    searchable_text.extend(
        message["content"]
        for message in get_messages(conversation["id"], current_user_id)
    )
    return normalized_term in " ".join(searchable_text).casefold()


def format_date(value):
    try:
        date = datetime.fromisoformat(str(value)).date()
        difference = (datetime.now().date() - date).days
        return "Today" if difference == 0 else "Yesterday" if difference == 1 else date.strftime("%b %d")
    except (ValueError, TypeError):
        return ""


def load_conversation(conversation_id):
    conversation = get_conversation(conversation_id, current_user_id)
    st.session_state.conversation_id = conversation_id
    st.session_state.messages = [
        {"role": row["role"], "content": row["content"], "document_name": row["document_name"]}
        for row in get_messages(conversation_id, current_user_id)
    ]
    st.session_state.document_name = conversation["document_name"] if conversation else None
    st.session_state.document_text = conversation["document_text"] if conversation else None


def start_new_conversation():
    load_conversation(create_conversation(current_user_id))


def build_prompt(user_message):
    parts = ["You are Nova, a helpful, friendly and professional AI assistant. Give clear and accurate answers."]
    for message in st.session_state.messages[:-1][-10:]:
        speaker = "User" if message["role"] == "user" else "Assistant"
        parts.append(f"{speaker}: {message['content']}")
    if st.session_state.document_text:
        parts.extend(["Current attachment content:", st.session_state.document_text[:50000]])
    parts.append(f"Current user message: {user_message}")
    return "\n".join(parts)


if "conversation_id" not in st.session_state:
    conversations = get_conversations(current_user_id)
    load_conversation(conversations[0]["id"] if conversations else create_conversation(current_user_id))


# ==================================================
# SIDEBAR
# ==================================================

with st.sidebar:
    brand_column, search_column = st.columns([5, 1], gap="small", vertical_alignment="center")
    with brand_column:
        st.markdown('<div class="brand-row"><div class="brand-icon">✣</div><div class="brand-name">Nova AI</div></div>', unsafe_allow_html=True)
    with search_column:
        if st.button("", icon=":material/search:", help="Search conversations", key="toggle_search"):
            st.session_state.search_open = not st.session_state.get("search_open", False)
            st.rerun()

    if st.button("New chat", icon=":material/edit_square:", help="Start a new chat", key="new_chat"):
        start_new_conversation()
        st.rerun()

    search = ""
    if st.session_state.get("search_open", False):
        search = st.text_input("Search conversations", placeholder="Search conversations", label_visibility="collapsed", key="conversation_search")

    with st.container(key="conversation-history"):
        st.markdown('<div class="section-label conversations-label">Recent conversations</div>', unsafe_allow_html=True)
        conversations = get_conversations(current_user_id)
        if search:
            conversations = [
                item for item in conversations
                if conversation_matches_search(item, search)
            ]

        for conversation in conversations:
            conversation_id = conversation["id"]
            selected = conversation_id == st.session_state.conversation_id
            chat_column, delete_column = st.columns([6, 1], vertical_alignment="center")
            with chat_column:
                label = f"💬 {conversation['title']} · {format_date(conversation['created_at'])}"
                if st.button(label, key=f"open_{conversation_id}", use_container_width=True, type="primary" if selected else "secondary"):
                    load_conversation(conversation_id)
                    st.rerun()
            with delete_column:
                if st.button("×", key=f"delete_{conversation_id}", help="Delete conversation"):
                    delete_conversation(conversation_id, current_user_id)
                    if selected:
                        remaining = get_conversations(current_user_id)
                        load_conversation(remaining[0]["id"] if remaining else create_conversation(current_user_id))
                    st.rerun()

    st.divider()
    if st.session_state.document_name:
        name = escape_html(st.session_state.document_name)
        st.markdown(f'<div class="attachment-card"><div class="attachment-label">Attached file</div><div class="attachment-name">📎 {name}</div></div>', unsafe_allow_html=True)
        if st.button("Remove attachment", use_container_width=True, key="remove_attachment"):
            remove_document(st.session_state.conversation_id, current_user_id)
            st.session_state.document_name = st.session_state.document_text = None
            st.rerun()

    if st.session_state.pop("reset_clear_confirmation", False):
        st.session_state.pop("confirm_clear_conversation", None)

    if st.button("Clear conversation", use_container_width=True, key="clear_conversation"):
        st.session_state.pop("confirm_clear_conversation", None)
        st.session_state.show_clear_confirmation = True

    if st.session_state.get("show_clear_confirmation", False):
        st.markdown('<div class="clear-confirmation">This will remove all messages from this conversation.</div>', unsafe_allow_html=True)
        confirm_clear = st.checkbox("Confirm clear", key="confirm_clear_conversation")
        if st.button("Confirm", use_container_width=True, disabled=not confirm_clear, key="confirm_clear_action"):
            clear_conversation(st.session_state.conversation_id, current_user_id)
            load_conversation(st.session_state.conversation_id)
            st.session_state.show_clear_confirmation = False
            st.session_state.reset_clear_confirmation = True
            st.rerun()

    with st.container(key="sidebar-profile"):
        st.markdown(f'<div class="profile-name">👤 {escape_html(st.session_state.user["name"])}</div>', unsafe_allow_html=True)
        st.caption(st.session_state.user["email"])
        if st.button("↪  Logout", use_container_width=True, key="logout"):
            logout_user()
            st.rerun()


# ==================================================
# MAIN CHAT
# ==================================================

current = get_conversation(st.session_state.conversation_id, current_user_id)
title = escape_html(current["title"] if current else "New conversation")
st.markdown(
    f'<div class="nova-header"><div><p class="nova-header-title">{title}</p>'
    '<div class="nova-header-status">Nova 2.0 · Personal assistant</div></div>'
    '<div class="ready-badge" aria-label="Ready"><span class="ready-dot"></span></div></div>',
    unsafe_allow_html=True,
)

suggested_prompt = None
if not st.session_state.messages:
    st.markdown(
        '<div class="welcome-section"><div class="nova-logo">✣</div><div class="meet-nova">Meet Nova</div>'
        '<h1 class="welcome-title">How can I help you today?</h1>'
        '<p class="welcome-description">Ask a question, explore an idea, or attach a document or image.</p></div>',
        unsafe_allow_html=True,
    )
    left, right = st.columns(2)
    with left:
        if st.button("💡 Brainstorm ideas\n\nPlan a creative product launch", use_container_width=True):
            suggested_prompt = "Help me brainstorm ideas for a creative product launch."
        if st.button("✨ Improve my writing\n\nMake my writing clearer", use_container_width=True):
            suggested_prompt = "Help me improve my writing and make it clearer."
    with right:
        if st.button("📄 Summarize a document\n\nPull out the important points", use_container_width=True):
            if st.session_state.document_text:
                suggested_prompt = "Summarize the attached file and list its important points."
            else:
                st.warning("Attach a document first.")
        if st.button("💬 Learn something\n\nExplain a complex topic simply", use_container_width=True):
            suggested_prompt = "Explain artificial intelligence in simple terms."

for message in st.session_state.messages:
    with st.chat_message(message["role"], avatar="👤" if message["role"] == "user" else "✨"):
        st.markdown(message["content"])
        if message["role"] == "user" and message.get("document_name"):
            st.caption(f"📎 {message['document_name']}")

submission = st.chat_input(
    "Message Nova or attach a file...",
    accept_file=True,
    file_type=["pdf", "docx", "xlsx", "pptx", "png", "jpg", "jpeg", "webp"],
)

user_message = suggested_prompt or (submission.text.strip() if submission else "")
uploaded_files = submission.files if submission and not suggested_prompt else []

if user_message or uploaded_files:
    uploaded_name = None
    if uploaded_files:
        uploaded_file = uploaded_files[0]
        try:
            with st.spinner(f"Reading {uploaded_file.name}..."):
                extracted = extract_uploaded_file(uploaded_file)
            uploaded_name = uploaded_file.name
            st.session_state.document_name = uploaded_name
            st.session_state.document_text = extracted
            save_document(st.session_state.conversation_id, current_user_id, uploaded_name, extracted)
        except Exception as error:
            st.error(f"Could not read the file: {error}")
            st.stop()

    if not user_message and uploaded_name:
        extension = uploaded_name.rsplit(".", 1)[-1].lower()
        user_message = "Explain what is shown in this image." if extension in {"png", "jpg", "jpeg", "webp"} else "Summarize the attached document."

    if not st.session_state.messages:
        update_conversation_title(st.session_state.conversation_id, current_user_id, make_title(user_message))

    st.session_state.messages.append({"role": "user", "content": user_message, "document_name": uploaded_name})
    save_message(st.session_state.conversation_id, current_user_id, "user", user_message, uploaded_name)

    with st.chat_message("user", avatar="👤"):
        st.markdown(user_message)
        if uploaded_name:
            st.caption(f"📎 {uploaded_name}")

    with st.chat_message("assistant", avatar="✨"):
        if not client:
            assistant_message = "GEMINI_API_KEY was not found in the .env file."
            st.error(assistant_message)
        else:
            try:
                with st.spinner("Nova is thinking..."):
                    response = client.models.generate_content(model=model_name, contents=build_prompt(user_message))
                    assistant_message = response.text or "I could not generate a response."
                st.markdown(assistant_message)
            except Exception as error:
                assistant_message = f"Something went wrong: {error}"
                st.error(assistant_message)

    st.session_state.messages.append({"role": "assistant", "content": assistant_message, "document_name": None})
    save_message(st.session_state.conversation_id, current_user_id, "assistant", assistant_message)
    st.rerun()

st.markdown('<div class="disclaimer">Nova can make mistakes. Check important information.</div>', unsafe_allow_html=True)
