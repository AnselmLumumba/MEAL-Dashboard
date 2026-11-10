"""One-time account activation codes."""

from __future__ import annotations

import hashlib
import re
import secrets


_CODE_LENGTH = 6
_HASH_ITERATIONS = 200_000
_CODE_PATTERN = re.compile(r"^[0-9]{6}$")


def create_activation_code() -> str:
    return f"{secrets.randbelow(10**_CODE_LENGTH):0{_CODE_LENGTH}d}"


def hash_activation_code(code: str) -> str:
    if not _CODE_PATTERN.fullmatch(code):
        raise ValueError("Activation codes must contain six digits.")
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        code.encode("ascii"),
        salt,
        _HASH_ITERATIONS,
    )
    return f"pbkdf2_sha256${_HASH_ITERATIONS}${salt.hex()}${digest.hex()}"


def verify_activation_code(code: str, saved_hash: str) -> bool:
    if not _CODE_PATTERN.fullmatch(code):
        return False
    try:
        algorithm, iterations_text, salt_hex, digest_hex = saved_hash.split("$", 3)
        iterations = int(iterations_text)
        if algorithm != "pbkdf2_sha256" or not 100_000 <= iterations <= 2_000_000:
            return False
        expected = hashlib.pbkdf2_hmac(
            "sha256",
            code.encode("ascii"),
            bytes.fromhex(salt_hex),
            iterations,
        ).hex()
    except (ValueError, TypeError):
        return False
    return secrets.compare_digest(expected, digest_hex)
