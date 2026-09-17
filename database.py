import sqlite3

from auth import hash_password, verify_password


DATABASE_NAME = "chatbot.db"


def connect_database():
    connection = sqlite3.connect(DATABASE_NAME)

    connection.row_factory = sqlite3.Row

    # Enable SQLite foreign-key support
    connection.execute("PRAGMA foreign_keys = ON")

    return connection


def create_tables():
    connection = connect_database()
    cursor = connection.cursor()

    # Users
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    # Conversations
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS conversations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            title TEXT NOT NULL DEFAULT 'New conversation',
            document_name TEXT,
            document_text TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id)
                REFERENCES users(id)
                ON DELETE CASCADE
        )
        """
    )

    # Messages
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            conversation_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            document_name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (conversation_id)
                REFERENCES conversations(id)
                ON DELETE CASCADE
        )
        """
    )

    # Upgrade an older conversations table
    conversation_columns = {
        column["name"]
        for column in cursor.execute(
            "PRAGMA table_info(conversations)"
        ).fetchall()
    }

    if "user_id" not in conversation_columns:
        cursor.execute(
            """
            ALTER TABLE conversations
            ADD COLUMN user_id INTEGER
            """
        )

    if "document_name" not in conversation_columns:
        cursor.execute(
            """
            ALTER TABLE conversations
            ADD COLUMN document_name TEXT
            """
        )

    if "document_text" not in conversation_columns:
        cursor.execute(
            """
            ALTER TABLE conversations
            ADD COLUMN document_text TEXT
            """
        )

    if "updated_at" not in conversation_columns:
        cursor.execute(
            """
            ALTER TABLE conversations
            ADD COLUMN updated_at TIMESTAMP
            """
        )

    # Upgrade an older messages table
    message_columns = {
        column["name"]
        for column in cursor.execute(
            "PRAGMA table_info(messages)"
        ).fetchall()
    }

    if "document_name" not in message_columns:
        cursor.execute(
            """
            ALTER TABLE messages
            ADD COLUMN document_name TEXT
            """
        )

    connection.commit()
    connection.close()


# ==================================================
# USER FUNCTIONS
# ==================================================

def create_user(name, email, password):
    connection = connect_database()
    cursor = connection.cursor()

    clean_name = name.strip()
    clean_email = email.strip().lower()
    password_hash = hash_password(password)

    try:
        cursor.execute(
            """
            INSERT INTO users (
                name,
                email,
                password_hash
            )
            VALUES (?, ?, ?)
            """,
            (
                clean_name,
                clean_email,
                password_hash,
            ),
        )

        user_id = cursor.lastrowid

        connection.commit()
        connection.close()

        return {
            "success": True,
            "user_id": user_id,
            "message": "Account created successfully.",
        }

    except sqlite3.IntegrityError:
        connection.close()

        return {
            "success": False,
            "user_id": None,
            "message": (
                "An account already exists with this email."
            ),
        }


def authenticate_user(email, password):
    connection = connect_database()
    cursor = connection.cursor()

    clean_email = email.strip().lower()

    cursor.execute(
        """
        SELECT id, name, email, password_hash
        FROM users
        WHERE email = ?
        """,
        (clean_email,),
    )

    user = cursor.fetchone()
    connection.close()

    if not user:
        return None

    if not verify_password(
        password,
        user["password_hash"],
    ):
        return None

    return {
        "id": user["id"],
        "name": user["name"],
        "email": user["email"],
    }


def get_user(user_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT id, name, email, created_at
        FROM users
        WHERE id = ?
        """,
        (user_id,),
    )

    user = cursor.fetchone()
    connection.close()

    return user


# ==================================================
# CONVERSATION FUNCTIONS
# ==================================================

def create_conversation(
    user_id,
    title="New conversation",
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        INSERT INTO conversations (
            user_id,
            title
        )
        VALUES (?, ?)
        """,
        (
            user_id,
            title,
        ),
    )

    conversation_id = cursor.lastrowid

    connection.commit()
    connection.close()

    return conversation_id


def get_conversations(user_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            id,
            title,
            document_name,
            created_at,
            updated_at
        FROM conversations
        WHERE user_id = ?
        ORDER BY
            COALESCE(updated_at, created_at) DESC,
            id DESC
        """,
        (user_id,),
    )

    conversations = cursor.fetchall()
    connection.close()

    return conversations


def get_conversation(
    conversation_id,
    user_id,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            id,
            user_id,
            title,
            document_name,
            document_text,
            created_at,
            updated_at
        FROM conversations
        WHERE id = ?
        AND user_id = ?
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    conversation = cursor.fetchone()
    connection.close()

    return conversation


def update_conversation_title(
    conversation_id,
    user_id,
    title,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE conversations
        SET
            title = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        AND user_id = ?
        """,
        (
            title,
            conversation_id,
            user_id,
        ),
    )

    connection.commit()
    connection.close()


def save_document(
    conversation_id,
    user_id,
    document_name,
    document_text,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE conversations
        SET
            document_name = ?,
            document_text = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        AND user_id = ?
        """,
        (
            document_name,
            document_text,
            conversation_id,
            user_id,
        ),
    )

    connection.commit()
    connection.close()


def remove_document(
    conversation_id,
    user_id,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE conversations
        SET
            document_name = NULL,
            document_text = NULL,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        AND user_id = ?
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    connection.commit()
    connection.close()


# ==================================================
# MESSAGE FUNCTIONS
# ==================================================

def save_message(
    conversation_id,
    user_id,
    role,
    content,
    document_name=None,
):
    connection = connect_database()
    cursor = connection.cursor()

    # Confirm that the conversation belongs to the user
    cursor.execute(
        """
        SELECT id
        FROM conversations
        WHERE id = ?
        AND user_id = ?
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    conversation = cursor.fetchone()

    if not conversation:
        connection.close()

        raise ValueError(
            "Conversation does not belong to this user."
        )

    cursor.execute(
        """
        INSERT INTO messages (
            conversation_id,
            role,
            content,
            document_name
        )
        VALUES (?, ?, ?, ?)
        """,
        (
            conversation_id,
            role,
            content,
            document_name,
        ),
    )

    cursor.execute(
        """
        UPDATE conversations
        SET updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        AND user_id = ?
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    connection.commit()
    connection.close()


def get_messages(
    conversation_id,
    user_id,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT
            messages.role,
            messages.content,
            messages.document_name,
            messages.created_at
        FROM messages
        INNER JOIN conversations
            ON messages.conversation_id
            = conversations.id
        WHERE messages.conversation_id = ?
        AND conversations.user_id = ?
        ORDER BY messages.id ASC
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    messages = cursor.fetchall()
    connection.close()

    return messages


def clear_conversation(
    conversation_id,
    user_id,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        DELETE FROM messages
        WHERE conversation_id IN (
            SELECT id
            FROM conversations
            WHERE id = ?
            AND user_id = ?
        )
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    cursor.execute(
        """
        UPDATE conversations
        SET
            title = 'New conversation',
            document_name = NULL,
            document_text = NULL,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        AND user_id = ?
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    connection.commit()
    connection.close()


def delete_conversation(
    conversation_id,
    user_id,
):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        DELETE FROM messages
        WHERE conversation_id IN (
            SELECT id
            FROM conversations
            WHERE id = ?
            AND user_id = ?
        )
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    cursor.execute(
        """
        DELETE FROM conversations
        WHERE id = ?
        AND user_id = ?
        """,
        (
            conversation_id,
            user_id,
        ),
    )

    connection.commit()
    connection.close()