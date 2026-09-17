import hashlib
import hmac
import os
import re


PBKDF2_ITERATIONS = 600_000


def hash_password(password):
    salt = os.urandom(32)

    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PBKDF2_ITERATIONS,
    )

    return f"{salt.hex()}:{password_hash.hex()}"


def verify_password(password, stored_password):
    try:
        salt_hex, stored_hash_hex = (
            stored_password.split(":", 1)
        )

        salt = bytes.fromhex(salt_hex)
        stored_hash = bytes.fromhex(stored_hash_hex)

        supplied_hash = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            PBKDF2_ITERATIONS,
        )

        return hmac.compare_digest(
            stored_hash,
            supplied_hash,
        )

    except (ValueError, TypeError):
        return False


def validate_name(name):
    cleaned_name = name.strip()

    if len(cleaned_name) < 2:
        return False, "Name must contain at least 2 characters."

    if len(cleaned_name) > 50:
        return False, "Name cannot exceed 50 characters."

    return True, ""


def validate_email(email):
    cleaned_email = email.strip().lower()

    if not cleaned_email:
        return False, "Email address is required."

    if cleaned_email.count("@") != 1:
        return False, "Enter a valid email address."

    username, domain = cleaned_email.split("@")

    if not username:
        return False, "Enter a valid email address."

    if "." not in domain:
        return False, "Enter a valid email address."

    if domain.startswith(".") or domain.endswith("."):
        return False, "Enter a valid email address."

    return True, ""


def validate_password(password):
    if len(password) < 8:
        return (
            False,
            "Password must contain at least 8 characters.",
        )

    if not any(character.isupper() for character in password):
        return (
            False,
            "Password must contain an uppercase letter.",
        )

    if not any(character.islower() for character in password):
        return (
            False,
            "Password must contain a lowercase letter.",
        )

    if not any(character.isdigit() for character in password):
        return (
            False,
            "Password must contain a number.",
        )

    return True, ""