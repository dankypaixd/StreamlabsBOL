"""Memoria de contactos confirmados: se aprende en segundo plano y ayuda a leer mejor.

Que guarda (solo lo que la persona confirmo como correcto):
* correo -> nombre y telefonos vistos con ese correo
* dominios de correo conocidos
* codigo postal -> estado

Para que sirve: corregir confusiones tipicas del OCR (l/i/1, o/0, rn/m...) contra algo ya
verificado, y pasar a verde los contactos que repiten exactamente algo confirmado.

Donde se guarda: en el propio repositorio de GitHub, en una rama aparte (`bol-memory`) para
que cada guardado NO redespliegue la app. Como son datos de clientes y el repositorio puede
ser publico, el archivo va cifrado (Fernet). Sin token de GitHub usa un archivo local.
"""
from __future__ import annotations

import base64
import difflib
import hashlib
import json
import os
import re
import threading
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Mapping, Protocol

DEFAULT_REPO = "dankypaixd/StreamlabsBOL"
DEFAULT_BRANCH = "bol-memory"
DEFAULT_PATH = "memory/contacts.bin"
API = "https://api.github.com"

# Letras/digitos que el OCR confunde entre si. Se usa solo para decidir si dos textos
# son "el mismo con una confusion de OCR", no para corregir cualquier parecido.
_FOLD = str.maketrans({"i": "1", "l": "1", "|": "1", "!": "1", "o": "0", "s": "5", "b": "8", "z": "2", "v": "u"})


def fold(text: str) -> str:
    return (text or "").lower().replace("rn", "m").translate(_FOLD)


def _empty() -> dict[str, Any]:
    return {"version": 1, "contacts": {}, "domains": {}, "zips": {}}


# ------------------------------------------------------------------ almacenes


class Store(Protocol):
    def load(self) -> bytes | None: ...
    def save(self, data: bytes) -> None: ...


class Conflict(RuntimeError):
    """Alguien guardo antes que nosotros (sha distinto)."""


class LocalStore:
    def __init__(self, path: str | os.PathLike[str] = ".bol_memory/contacts.bin") -> None:
        self.path = Path(path)

    def load(self) -> bytes | None:
        try:
            return self.path.read_bytes()
        except FileNotFoundError:
            return None

    def save(self, data: bytes) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_bytes(data)
        tmp.replace(self.path)


class GitHubStore:
    """Archivo en una rama del repositorio, via la API de contenidos de GitHub."""

    def __init__(
        self,
        token: str,
        repo: str = DEFAULT_REPO,
        branch: str = DEFAULT_BRANCH,
        path: str = DEFAULT_PATH,
        session: Any | None = None,
    ) -> None:
        if session is None:
            import requests

            session = requests.Session()
        self.repo, self.branch, self.path = repo, branch, path
        self.http = session
        self.http.headers.update(
            {
                "Authorization": f"Bearer {token}",
                "Accept": "application/vnd.github+json",
                "X-GitHub-Api-Version": "2022-11-28",
                "User-Agent": "bol-memory",
            }
        )
        self._sha: str | None = None

    def _url(self, suffix: str) -> str:
        return f"{API}/repos/{self.repo}{suffix}"

    def load(self) -> bytes | None:
        response = self.http.get(self._url(f"/contents/{self.path}"), params={"ref": self.branch}, timeout=20)
        if response.status_code == 404:
            self._sha = None
            return None
        response.raise_for_status()
        body = response.json()
        self._sha = body.get("sha")
        return base64.b64decode(body.get("content", ""))

    def _ensure_branch(self) -> None:
        ref = self.http.get(self._url(f"/git/ref/heads/{self.branch}"), timeout=20)
        if ref.status_code == 200:
            return
        if ref.status_code != 404:
            ref.raise_for_status()
        info = self.http.get(self._url(""), timeout=20)
        info.raise_for_status()
        default = info.json()["default_branch"]
        base = self.http.get(self._url(f"/git/ref/heads/{default}"), timeout=20)
        base.raise_for_status()
        created = self.http.post(
            self._url("/git/refs"),
            json={"ref": f"refs/heads/{self.branch}", "sha": base.json()["object"]["sha"]},
            timeout=20,
        )
        if created.status_code not in (200, 201, 422):  # 422: la rama ya existe (carrera)
            created.raise_for_status()

    def save(self, data: bytes) -> None:
        self._ensure_branch()
        payload: dict[str, Any] = {
            "message": "bol-memory: actualizar datos aprendidos",
            "content": base64.b64encode(data).decode("ascii"),
            "branch": self.branch,
        }
        if self._sha:
            payload["sha"] = self._sha
        response = self.http.put(self._url(f"/contents/{self.path}"), json=payload, timeout=30)
        if response.status_code in (409, 422):
            raise Conflict(response.text[:200])
        response.raise_for_status()
        self._sha = response.json().get("content", {}).get("sha", self._sha)


# ------------------------------------------------------------------- cifrado


def _fernet(secret: str):
    from cryptography.fernet import Fernet

    return Fernet(base64.urlsafe_b64encode(hashlib.sha256(secret.encode("utf-8")).digest()))


def encode(data: dict[str, Any], secret: str | None) -> bytes:
    raw = json.dumps(data, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if secret:
        return json.dumps({"v": 1, "enc": "fernet", "data": _fernet(secret).encrypt(raw).decode("ascii")}).encode()
    return json.dumps({"v": 1, "enc": "none", "data": raw.decode("utf-8")}).encode()


def decode(blob: bytes | None, secret: str | None) -> dict[str, Any]:
    if not blob:
        return _empty()
    try:
        envelope = json.loads(blob.decode("utf-8"))
        if envelope.get("enc") == "fernet":
            if not secret:
                return _empty()
            raw = _fernet(secret).decrypt(envelope["data"].encode("ascii"))
        else:
            raw = envelope["data"].encode("utf-8")
        data = json.loads(raw)
    except Exception:  # archivo daniado o llave distinta: se empieza de cero
        return _empty()
    for key, default in _empty().items():
        data.setdefault(key, default)
    return data


# ------------------------------------------------------------------- memoria


def _top(counter: Mapping[str, int]) -> str:
    return max(counter, key=lambda k: (counter[k], k)) if counter else ""


def _merge_counts(a: dict[str, int], b: dict[str, int]) -> dict[str, int]:
    return {k: max(a.get(k, 0), b.get(k, 0)) for k in set(a) | set(b)}


def merge_data(mine: dict[str, Any], theirs: dict[str, Any]) -> dict[str, Any]:
    """Une dos versiones de la memoria (se usa al resolver un conflicto de guardado)."""
    out = _empty()
    for email in set(mine["contacts"]) | set(theirs["contacts"]):
        a, b = mine["contacts"].get(email, {}), theirs["contacts"].get(email, {})
        out["contacts"][email] = {
            "names": _merge_counts(a.get("names", {}), b.get("names", {})),
            "phones": _merge_counts(a.get("phones", {}), b.get("phones", {})),
            "count": max(a.get("count", 0), b.get("count", 0)),
            "last": max(a.get("last", ""), b.get("last", "")),
        }
    out["domains"] = _merge_counts(mine["domains"], theirs["domains"])
    for zip_code in set(mine["zips"]) | set(theirs["zips"]):
        out["zips"][zip_code] = _merge_counts(mine["zips"].get(zip_code, {}), theirs["zips"].get(zip_code, {}))
    return out


class ContactMemory:
    def __init__(self, store: Store | None = None, secret: str | None = None, label: str = "local") -> None:
        self.store = store or LocalStore()
        self.secret = secret
        self.label = label
        self.last_error = ""
        self._lock = threading.RLock()
        self.data = _empty()
        self.reload()

    # -- carga / guardado --------------------------------------------------

    def reload(self) -> None:
        with self._lock:
            try:
                self.data = decode(self.store.load(), self.secret)
                self.last_error = ""
            except Exception as exc:  # sin red, token vencido...
                self.last_error = f"load: {exc}"

    def save(self, attempts: int = 3) -> bool:
        with self._lock:
            for _ in range(attempts):
                try:
                    self.store.save(encode(self.data, self.secret))
                    self.last_error = ""
                    return True
                except Conflict:
                    try:
                        remote = decode(self.store.load(), self.secret)
                    except Exception as exc:
                        self.last_error = f"save: {exc}"
                        return False
                    self.data = merge_data(self.data, remote)
                except Exception as exc:
                    self.last_error = f"save: {exc}"
                    return False
            self.last_error = "save: conflict"
            return False

    def save_async(self) -> threading.Thread:
        thread = threading.Thread(target=self.save, name="bol-memory-save", daemon=True)
        thread.start()
        return thread

    # -- consultas ---------------------------------------------------------

    def lookup(self, email: str) -> dict[str, Any] | None:
        entry = self.data["contacts"].get((email or "").strip().lower())
        if not entry:
            return None
        return {"name": _top(entry.get("names", {})), "phones": dict(entry.get("phones", {})), "count": entry.get("count", 0)}

    def closest_email(self, email: str, cutoff: float = 0.92) -> str | None:
        """Correo ya confirmado que solo difiere por confusiones tipicas de OCR."""
        target = (email or "").strip().lower()
        if not target or target in self.data["contacts"]:
            return None
        folded = fold(target)
        matches = [known for known in self.data["contacts"] if fold(known) == folded]
        if len(matches) == 1:
            return matches[0]
        if matches:
            return max(matches, key=lambda k: difflib.SequenceMatcher(None, target, k).ratio())
        return None

    def closest_domain(self, domain: str, cutoff: float = 0.9) -> str | None:
        domain = (domain or "").strip().lower()
        if not domain or domain in self.data["domains"]:
            return None
        folded = fold(domain)
        matches = [known for known in self.data["domains"] if fold(known) == folded]
        return matches[0] if len(matches) == 1 else None

    def state_for_zip(self, zip_code: str) -> str:
        return _top(self.data["zips"].get((zip_code or "").strip()[:5], {}))

    # -- aprendizaje -------------------------------------------------------

    def learn(self, rows: Iterable[Mapping[str, str]]) -> int:
        """Aprende contactos confirmados. Cada fila: name, email, phone, state, zip_code."""
        learned = 0
        today = date.today().isoformat()
        with self._lock:
            for row in rows:
                email = (row.get("email") or "").strip().lower()
                name = re.sub(r"\s+", " ", row.get("name") or "").strip()
                phone = row.get("phone") or ""
                if not (email and name and phone):
                    continue
                entry = self.data["contacts"].setdefault(email, {"names": {}, "phones": {}, "count": 0, "last": today})
                entry["names"][name] = entry["names"].get(name, 0) + 1
                entry["phones"][phone] = entry["phones"].get(phone, 0) + 1
                entry["count"] = entry.get("count", 0) + 1
                entry["last"] = today
                domain = email.partition("@")[2]
                if domain:
                    self.data["domains"][domain] = self.data["domains"].get(domain, 0) + 1
                zip5 = (row.get("zip_code") or "").strip()[:5]
                state = (row.get("state") or "").strip().upper()
                if zip5.isdigit() and len(zip5) == 5 and state:
                    bucket = self.data["zips"].setdefault(zip5, {})
                    bucket[state] = bucket.get(state, 0) + 1
                learned += 1
        return learned


def from_secrets(secrets: Mapping[str, Any] | None, local_path: str | os.PathLike[str] = ".bol_memory/contacts.bin") -> ContactMemory:
    """Crea la memoria con GitHub si hay GITHUB_TOKEN en los secretos; si no, archivo local."""
    secrets = secrets or {}
    token = str(secrets.get("GITHUB_TOKEN", "") or "")
    if not token:
        return ContactMemory(LocalStore(local_path), None, "local")
    store = GitHubStore(
        token,
        repo=str(secrets.get("GITHUB_REPO", DEFAULT_REPO) or DEFAULT_REPO),
        branch=str(secrets.get("MEMORY_BRANCH", DEFAULT_BRANCH) or DEFAULT_BRANCH),
    )
    # Cifrado con MEMORY_KEY; si no existe, se deriva del token (un token nuevo = empezar de cero).
    secret = str(secrets.get("MEMORY_KEY", "") or token)
    return ContactMemory(store, secret, "github")
