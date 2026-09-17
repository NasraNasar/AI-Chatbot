import os
from io import BytesIO

import streamlit as st
from dotenv import load_dotenv
from google import genai
from pypdf import PdfReader

from database import (
    clear_conversation,
    create_conversation,
    create_tables,
    delete_conversation,
    get_conversation,
    get_conversations,
    get_messages,
    remove_document,
    save_document,
    save_message,
    update_conversation_title,
)


# ---------------------------------------------------
# Page configuration
# ---------------------------------------------------

st.set_page_config(
    page_title="Nova AI",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

load_dotenv()
create_tables()


# ---------------------------------------------------
# Styling
# ---------------------------------------------------

st.markdown(
    """
    <style>
        .stApp {
            background:
                radial-gradient(
                    circle at top right,
                    #eeeaff 0%,
                    #f7f8ff 35%,
                    #ffffff 75%
                );
        }

        [data-testid="stSidebar"] {
            background:
                linear-gradient(
                    180deg,
                    #0f172a 0%,
                    #17173b 55%,
                    #21194f 100%
                );
        }

        [data-testid="stSidebar"] * {
            color: white;
        }

        [data-testid="stSidebar"] button {
            border-radius: 12px;
        }

        [data-testid="stChatMessage"] {
            background: rgba(255, 255, 255, 0.90);
            border: 1px solid #e5e7eb;
            border-radius: 20px;
            padding: 12px 18px;
            margin-bottom: 14px;
            box-shadow: 0 8px 25px rgba(48, 46, 129, 0.06);
        }

        [data-testid="stChatInput"] {
            border-radius: 18px;
            border: 1px solid #8b5cf6;
            box-shadow: 0 8px 30px rgba(124, 58, 237, 0.12);
        }

        .main-title {
            font-size: 2.2rem;
            font-weight: 800;
            margin-bottom: 0;
        }

        .subtitle {
            color: #6b7280;
            margin-top: 4px;
            margin-bottom: 25px;
        }

        .document-box {
            padding: 14px;
            border-radius: 14px;
            background: rgba(99, 102, 241, 0.18);
            margin-bottom: 12px;
        }

        .small-text {
            color: #cbd5e1;
            font-size: 0.84rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------
# Helper functions
# ---------------------------------------------------

def extract_pdf_text(uploaded_file):
    pdf_bytes = uploaded_file.getvalue()
    reader = PdfReader(BytesIO(pdf_bytes))

    extracted_pages = []

    for page in reader.pages:
        page_text = page.extract_text()

        if page_text:
            extracted_pages.append(page_text)

    return "\n\n".join(extracted_pages).strip()


def make_conversation_title(message):
    cleaned_message = " ".join(message.split())

    if not cleaned_message:
        return "New conversation"

    if len(cleaned_message) > 35:
        return cleaned_message[:35] + "..."

    return cleaned_message


def load_conversation(conversation_id):
    conversation = get_conversation(conversation_id)
    database_messages = get_messages(conversation_id)

    st.session_state.conversation_id = conversation_id
    st.session_state.messages = []

    for message in database_messages:
        st.session_state.messages.append(
            {
                "role": message["role"],
                "content": message["content"],
                "document_name": message["document_name"],
            }
        )

    if conversation:
        st.session_state.document_name = (
            conversation["document_name"]
        )
        st.session_state.document_text = (
            conversation["document_text"]
        )
    else:
        st.session_state.document_name = None
        st.session_state.document_text = None


def start_new_conversation():
    new_id = create_conversation()
    load_conversation(new_id)


def build_prompt(user_message):
    prompt_sections = [
        (
            "You are Nova, a helpful and friendly AI assistant. "
            "Give clear and accurate answers."
        )
    ]

    # Add a limited amount of previous conversation
    recent_messages = st.session_state.messages[-10:]

    if recent_messages:
        prompt_sections.append("\nConversation history:")

        for message in recent_messages:
            speaker = (
                "User"
                if message["role"] == "user"
                else "Assistant"
            )

            prompt_sections.append(
                f"{speaker}: {message['content']}"
            )

    if st.session_state.document_text:
        # Limit the PDF text to avoid sending extremely large prompts
        pdf_text = st.session_state.document_text[:50000]

        prompt_sections.append(
            "\nAttached PDF document:"
        )
        prompt_sections.append(pdf_text)
        prompt_sections.append(
            "\nUse the PDF when it is relevant to the user's question."
        )

    prompt_sections.append(
        f"\nCurrent user message: {user_message}"
    )

    return "\n".join(prompt_sections)


# ---------------------------------------------------
# Initialize session state
# ---------------------------------------------------

if "conversation_id" not in st.session_state:
    existing_conversations = get_conversations()

    if existing_conversations:
        latest_conversation_id = existing_conversations[0]["id"]
    else:
        latest_conversation_id = create_conversation()

    load_conversation(latest_conversation_id)

if "messages" not in st.session_state:
    load_conversation(st.session_state.conversation_id)

if "document_name" not in st.session_state:
    st.session_state.document_name = None

if "document_text" not in st.session_state:
    st.session_state.document_text = None


# ---------------------------------------------------
# Gemini client
# ---------------------------------------------------

api_key = os.getenv("GEMINI_API_KEY")

if api_key:
    client = genai.Client(api_key=api_key)
else:
    client = None

# Keep your currently working Gemini model here.
# You can also add GEMINI_MODEL to the .env file.
model_name = os.getenv(
    "GEMINI_MODEL",
    "gemini-3.1-flash-lite",
)


# ---------------------------------------------------
# Sidebar
# ---------------------------------------------------

with st.sidebar:
    st.markdown("## ✨ Nova AI")
    st.markdown(
        '<p class="small-text">Your personal AI assistant</p>',
        unsafe_allow_html=True,
    )

    st.divider()

    if st.button(
        "＋ New chat",
        use_container_width=True,
        type="primary",
    ):
        start_new_conversation()
        st.rerun()

    st.markdown("### 💬 Previous chats")

    conversations = get_conversations()

    if not conversations:
        st.caption("No saved conversations")
    else:
        for conversation in conversations:
            conversation_id = conversation["id"]
            conversation_title = conversation["title"]

            button_type = (
                "primary"
                if conversation_id
                == st.session_state.conversation_id
                else "secondary"
            )

            chat_column, delete_column = st.columns(
                [5, 1],
                vertical_alignment="center",
            )

            # Open conversation
            with chat_column:
                if st.button(
                    f"💬 {conversation_title}",
                    key=f"open_chat_{conversation_id}",
                    use_container_width=True,
                    type=button_type,
                ):
                    load_conversation(conversation_id)
                    st.rerun()

            # Delete conversation
            with delete_column:
                if st.button(
                    "🗑️",
                    key=f"delete_chat_{conversation_id}",
                    help=f"Delete {conversation_title}",
                    use_container_width=True,
                ):
                    deleting_current_chat = (
                        conversation_id
                        == st.session_state.conversation_id
                    )

                    delete_conversation(conversation_id)

                    # If the open chat was deleted, open another chat
                    if deleting_current_chat:
                        remaining_conversations = get_conversations()

                        if remaining_conversations:
                            next_conversation_id = (
                                remaining_conversations[0]["id"]
                            )
                        else:
                            next_conversation_id = (
                                create_conversation()
                            )

                        load_conversation(next_conversation_id)

                    st.rerun()

    st.divider()
    st.markdown("### 📄 Current document")

    if st.session_state.document_name:
        st.markdown(
            f"""
            <div class="document-box">
                <strong>Attached PDF</strong><br>
                {st.session_state.document_name}
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button(
            "Remove PDF",
            use_container_width=True,
        ):
            remove_document(
                st.session_state.conversation_id
            )

            st.session_state.document_name = None
            st.session_state.document_text = None
            st.rerun()
    else:
        st.info(
            "No PDF attached. Use the attachment button "
            "beside the message box."
        )

    st.divider()

    if st.button(
        "🗑️ Clear current chat",
        use_container_width=True,
    ):
        clear_conversation(
            st.session_state.conversation_id
        )

        load_conversation(
            st.session_state.conversation_id
        )

        st.rerun()

    if st.button(
        "❌ Delete current chat",
        use_container_width=True,
    ):
        delete_conversation(
            st.session_state.conversation_id
        )

        remaining_conversations = get_conversations()

        if remaining_conversations:
            next_id = remaining_conversations[0]["id"]
        else:
            next_id = create_conversation()

        load_conversation(next_id)
        st.rerun()

    st.divider()
    st.caption("Powered by Gemini")


# ---------------------------------------------------
# Main interface
# ---------------------------------------------------

st.markdown(
    '<p class="main-title">✨ Nova AI Assistant</p>',
    unsafe_allow_html=True,
)

st.markdown(
    '<p class="subtitle">'
    'Ask a question or attach a PDF to begin.'
    '</p>',
    unsafe_allow_html=True,
)


# Welcome screen
if not st.session_state.messages:
    with st.container(border=True):
        st.markdown("### Welcome! 👋")
        st.write("You can ask questions such as:")

        st.markdown(
            """
            - Explain artificial intelligence simply.
            - Summarize the attached document.
            - What are the main points in this PDF?
            - Create five questions from the document.
            """
        )


# Display saved messages
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

        if (
            message["role"] == "user"
            and message.get("document_name")
        ):
            st.caption(
                f"📎 Attached: {message['document_name']}"
            )


# ---------------------------------------------------
# Chat input and PDF upload
# ---------------------------------------------------

submission = st.chat_input(
    "Type a message or attach a PDF...",
    accept_file=True,
    file_type=["pdf"],
)

if submission:
    user_message = submission.text.strip()
    uploaded_files = submission.files

    uploaded_document_name = None

    # Process the attached PDF
    if uploaded_files:
        uploaded_pdf = uploaded_files[0]

        try:
            pdf_text = extract_pdf_text(uploaded_pdf)

            if not pdf_text:
                st.error(
                    "No readable text was found in this PDF. "
                    "It may be a scanned image."
                )
                st.stop()

            uploaded_document_name = uploaded_pdf.name

            st.session_state.document_name = (
                uploaded_document_name
            )
            st.session_state.document_text = pdf_text

            save_document(
                st.session_state.conversation_id,
                uploaded_document_name,
                pdf_text,
            )

        except Exception as error:
            st.error(f"Could not read the PDF: {error}")
            st.stop()

    # Give the AI a default instruction when only a PDF is uploaded
    if not user_message and uploaded_document_name:
        user_message = "Please summarize this PDF."

    if not user_message:
        st.stop()

    # Update the title using the first user message
    if not st.session_state.messages:
        conversation_title = make_conversation_title(
            user_message
        )

        update_conversation_title(
            st.session_state.conversation_id,
            conversation_title,
        )

    # Save the user's message in session and SQLite
    user_data = {
        "role": "user",
        "content": user_message,
        "document_name": uploaded_document_name,
    }

    st.session_state.messages.append(user_data)

    save_message(
        st.session_state.conversation_id,
        "user",
        user_message,
        uploaded_document_name,
    )

    # Show the user's message immediately
    with st.chat_message("user"):
        st.markdown(user_message)

        if uploaded_document_name:
            st.caption(
                f"📎 Attached: {uploaded_document_name}"
            )

    # Generate the AI response
    with st.chat_message("assistant"):
        if not client:
            assistant_message = (
                "GEMINI_API_KEY was not found. "
                "Please add it to your .env file."
            )

            st.error(assistant_message)

        else:
            try:
                with st.spinner("Nova is thinking..."):
                    full_prompt = build_prompt(user_message)

                    response = client.models.generate_content(
                        model=model_name,
                        contents=full_prompt,
                    )

                    assistant_message = (
                        response.text
                        or "I could not generate a response."
                    )

                st.markdown(assistant_message)

            except Exception as error:
                assistant_message = (
                    f"Something went wrong: {error}"
                )

                st.error(assistant_message)

    # Save the assistant response
    assistant_data = {
        "role": "assistant",
        "content": assistant_message,
        "document_name": None,
    }

    st.session_state.messages.append(assistant_data)

    save_message(
        st.session_state.conversation_id,
        "assistant",
        assistant_message,
    )

    # Refresh sidebar document and conversation information
    st.rerun()