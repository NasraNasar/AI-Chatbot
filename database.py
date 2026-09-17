import sqlite3


DATABASE_NAME = "chatbot.db"


def connect_database():
    connection = sqlite3.connect(DATABASE_NAME)
    connection.row_factory = sqlite3.Row
    return connection


def create_tables():
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS conversations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL DEFAULT 'New conversation',
            document_name TEXT,
            document_text TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

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

    # Upgrade an older database if necessary
    conversation_columns = {
        column["name"]
        for column in cursor.execute(
            "PRAGMA table_info(conversations)"
        ).fetchall()
    }

    if "document_name" not in conversation_columns:
        cursor.execute(
            "ALTER TABLE conversations ADD COLUMN document_name TEXT"
        )

    if "document_text" not in conversation_columns:
        cursor.execute(
            "ALTER TABLE conversations ADD COLUMN document_text TEXT"
        )

    if "updated_at" not in conversation_columns:
        cursor.execute(
            "ALTER TABLE conversations ADD COLUMN updated_at TIMESTAMP"
        )

    message_columns = {
        column["name"]
        for column in cursor.execute(
            "PRAGMA table_info(messages)"
        ).fetchall()
    }

    if "document_name" not in message_columns:
        cursor.execute(
            "ALTER TABLE messages ADD COLUMN document_name TEXT"
        )

    connection.commit()
    connection.close()


def create_conversation(title="New conversation"):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        INSERT INTO conversations (title)
        VALUES (?)
        """,
        (title,),
    )

    conversation_id = cursor.lastrowid

    connection.commit()
    connection.close()

    return conversation_id


def get_conversations():
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT id, title, document_name, created_at
        FROM conversations
        ORDER BY COALESCE(updated_at, created_at) DESC, id DESC
        """
    )

    conversations = cursor.fetchall()
    connection.close()

    return conversations


def get_conversation(conversation_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT id, title, document_name, document_text, created_at
        FROM conversations
        WHERE id = ?
        """,
        (conversation_id,),
    )

    conversation = cursor.fetchone()
    connection.close()

    return conversation


def update_conversation_title(conversation_id, title):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE conversations
        SET title = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        """,
        (title, conversation_id),
    )

    connection.commit()
    connection.close()


def save_document(conversation_id, document_name, document_text):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE conversations
        SET document_name = ?,
            document_text = ?,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        """,
        (document_name, document_text, conversation_id),
    )

    connection.commit()
    connection.close()


def remove_document(conversation_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        UPDATE conversations
        SET document_name = NULL,
            document_text = NULL,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        """,
        (conversation_id,),
    )

    connection.commit()
    connection.close()


def save_message(
    conversation_id,
    role,
    content,
    document_name=None,
):
    connection = connect_database()
    cursor = connection.cursor()

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
        """,
        (conversation_id,),
    )

    connection.commit()
    connection.close()


def get_messages(conversation_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        SELECT role, content, document_name
        FROM messages
        WHERE conversation_id = ?
        ORDER BY id ASC
        """,
        (conversation_id,),
    )

    messages = cursor.fetchall()
    connection.close()

    return messages


def clear_conversation(conversation_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        DELETE FROM messages
        WHERE conversation_id = ?
        """,
        (conversation_id,),
    )

    cursor.execute(
        """
        UPDATE conversations
        SET title = 'New conversation',
            document_name = NULL,
            document_text = NULL,
            updated_at = CURRENT_TIMESTAMP
        WHERE id = ?
        """,
        (conversation_id,),
    )

    connection.commit()
    connection.close()


def delete_conversation(conversation_id):
    connection = connect_database()
    cursor = connection.cursor()

    cursor.execute(
        """
        DELETE FROM messages
        WHERE conversation_id = ?
        """,
        (conversation_id,),
    )

    cursor.execute(
        """
        DELETE FROM conversations
        WHERE id = ?
        """,
        (conversation_id,),
    )

    connection.commit()
    connection.close()