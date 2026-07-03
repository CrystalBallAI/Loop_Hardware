"""
crypt.py — at-rest encryption for the scoring IP (spec JSONs + recommendation
libraries) with transparent in-memory decryption.

Threat model (stated honestly to the user): an offline app can never be fully
un-reversible. This raises the cost. The data files ship as AES-256-GCM
ciphertext; the key is assembled at runtime from masked fragments. Once the app
is compiled with Nuitka, that assembly is machine code with no bytecode to
decompile — which is the actual protection. In dev (interpreted) it's only
obfuscation, by design.

Mechanism: encrypted files carry an 8-byte MAGIC header. `install_read_hook()`
wraps pathlib.Path.{read_text,read_bytes,open} so that ANY read of a
magic-prefixed file returns decrypted plaintext, and every other file passes
through untouched. This means none of the ~30 spec/library loader call sites in
the vendored pipelines need editing — they read "plaintext" transparently.
"""
from __future__ import annotations

import hashlib
import io
import os
import pathlib

MAGIC = b"CBMIENC1"
_NONCE = 12

# --- obfuscated key fragments (generated once; reassembled at runtime) ------
_K = [b'\x5e\x4f\x18\xec\x56\x3a\x7d\xe9', b'\x27\x32\xdc\x6c\x5a\xdb\x38\x38',
      b'\xb1\x24\x63\x4e\x17\x2e\xe2\x5a', b'\xb6\x2f\x46\xe2\xd7\x9f\x33\x9a']
_M = [b'\xb2\xa2\xea\x0e\x89\x49\xd6\xdd', b'\xed\x4c\x72\x36\x4d\x55\xf2\xac',
      b'\x5d\xb4\xcb\xad\x86\xf3\x98\x87', b'\x88\xaa\x8d\x83\x50\xc4\x37\x78']


def _key() -> bytes:
    raw = b"".join(bytes(a ^ b for a, b in zip(f, m)) for f, m in zip(_K, _M))
    # run through SHA-256 so the literal key never sits contiguously in memory
    return hashlib.sha256(raw).digest()


def _aesgcm():
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    return AESGCM(_key())


def is_encrypted(blob: bytes) -> bool:
    return blob[:len(MAGIC)] == MAGIC


def encrypt_bytes(plaintext: bytes) -> bytes:
    nonce = os.urandom(_NONCE)
    ct = _aesgcm().encrypt(nonce, plaintext, None)
    return MAGIC + nonce + ct


def decrypt_bytes(blob: bytes) -> bytes:
    nonce = blob[len(MAGIC):len(MAGIC) + _NONCE]
    ct = blob[len(MAGIC) + _NONCE:]
    return _aesgcm().decrypt(nonce, ct, None)


# --- transparent read hook --------------------------------------------------
# Capture the REAL Path.open. Everything reads raw bytes through this directly,
# never through Path.read_bytes/read_text — those re-dispatch to self.open,
# which is the hooked method, and would recurse infinitely.
_orig_open = pathlib.Path.open
_installed = False


def _raw_bytes(self) -> bytes:
    with _orig_open(self, "rb") as fh:
        return fh.read()


def encrypt_file(path: pathlib.Path) -> bool:
    """Encrypt a file in place (idempotent — skips already-encrypted)."""
    raw = _raw_bytes(path)
    if is_encrypted(raw):
        return False
    tmp = path.with_suffix(path.suffix + ".enc.tmp")
    with _orig_open(tmp, "wb") as fh:
        fh.write(encrypt_bytes(raw))
    os.replace(tmp, path)
    return True


def _read_bytes(self):
    data = _raw_bytes(self)
    return decrypt_bytes(data) if is_encrypted(data) else data


def _read_text(self, encoding=None, errors=None):
    data = _raw_bytes(self)
    if is_encrypted(data):
        return decrypt_bytes(data).decode(encoding or "utf-8", errors or "strict")
    with _orig_open(self, "r", encoding=encoding, errors=errors) as fh:
        return fh.read()


def _open(self, mode="r", buffering=-1, encoding=None,
          errors=None, newline=None):
    # only intercept pure read modes; writes/appends pass straight through
    if "r" in mode and set(mode) <= set("rbt") and "+" not in mode:
        try:
            with _orig_open(self, "rb") as fh:
                head = fh.read(len(MAGIC))
        except OSError:
            head = b""
        if head == MAGIC:
            pt = decrypt_bytes(_raw_bytes(self))
            if "b" in mode:
                return io.BytesIO(pt)
            return io.StringIO(pt.decode(encoding or "utf-8",
                                         errors or "strict"))
    return _orig_open(self, mode, buffering, encoding, errors, newline)


def install_read_hook() -> None:
    """Idempotently install transparent decryption on pathlib.Path reads.
    Call once per process that reads spec/library files (pipeline step children
    AND the parent process's adapter)."""
    global _installed
    if _installed:
        return
    pathlib.Path.read_bytes = _read_bytes
    pathlib.Path.read_text = _read_text
    pathlib.Path.open = _open
    _installed = True
