import base64
import json

import contact_memory as cm


class FakeResponse:
    def __init__(self, status=200, body=None):
        self.status_code, self._body, self.text = status, body or {}, json.dumps(body or {})

    def json(self):
        return self._body

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")


class FakeGitHub:
    """Simula la API de contenidos de GitHub (rama, archivo, sha y conflicto)."""

    def __init__(self):
        self.headers = {}
        self.branches = {"main": "sha-main"}
        self.file = None  # (bytes, sha)
        self.puts = 0
        self.conflict_once = False

    def get(self, url, params=None, timeout=None):
        if "/contents/" in url:
            if self.file is None or params["ref"] not in self.branches:
                return FakeResponse(404)
            return FakeResponse(200, {"sha": self.file[1], "content": base64.b64encode(self.file[0]).decode()})
        if "/git/ref/heads/" in url:
            name = url.rsplit("/", 1)[1]
            return FakeResponse(200, {"object": {"sha": self.branches[name]}}) if name in self.branches else FakeResponse(404)
        return FakeResponse(200, {"default_branch": "main"})

    def post(self, url, json=None, timeout=None):
        self.branches[json["ref"].rsplit("/", 1)[1]] = json["sha"]
        return FakeResponse(201)

    def put(self, url, json=None, timeout=None):
        if self.conflict_once:
            self.conflict_once = False
            return FakeResponse(409, {"message": "sha mismatch"})
        self.puts += 1
        self.file = (base64.b64decode(json["content"]), f"sha{self.puts}")
        return FakeResponse(200, {"content": {"sha": self.file[1]}})


ROW = {"name": "Jane Testwell", "email": "Jane.Testwell@Example.com", "phone": "1+15551230147", "state": "TX", "zip_code": "75201"}


def test_learn_and_lookup(tmp_path):
    memory = cm.ContactMemory(cm.LocalStore(tmp_path / "m.bin"))
    assert memory.learn([ROW, {"name": "", "email": "x@y.com", "phone": "1"}]) == 1
    entry = memory.lookup("jane.testwell@example.com")
    assert entry["name"] == "Jane Testwell" and "1+15551230147" in entry["phones"]
    assert memory.state_for_zip("75201") == "TX"


def test_ocr_confusions_only():
    memory = cm.ContactMemory(cm.LocalStore("/tmp/unused-bol-memory.bin"))
    memory.data = cm._empty()
    memory.learn([ROW])
    assert memory.closest_email("jane.testweil@example.com") == "jane.testwell@example.com"  # l <-> i
    assert memory.closest_email("jane.testwelk@example.com") is None  # no es una confusion de OCR
    assert memory.closest_domain("examp1e.com") == "example.com"


def test_encrypted_round_trip_and_wrong_key(tmp_path):
    store = cm.LocalStore(tmp_path / "m.bin")
    memory = cm.ContactMemory(store, secret="clave-uno")
    memory.learn([ROW])
    assert memory.save()
    raw = (tmp_path / "m.bin").read_bytes()
    assert b"Testwell" not in raw and b"example.com" not in raw  # cifrado
    assert cm.ContactMemory(store, secret="clave-uno").lookup("jane.testwell@example.com")
    assert cm.ContactMemory(store, secret="otra").lookup("jane.testwell@example.com") is None


def test_github_store_creates_branch_and_merges_conflicts():
    gh = FakeGitHub()
    memory = cm.ContactMemory(cm.GitHubStore("tok", session=gh), secret="tok", label="github")
    memory.learn([ROW])
    assert memory.save() and "bol-memory" in gh.branches and gh.puts == 1

    other = cm.ContactMemory(cm.GitHubStore("tok", session=gh), secret="tok")
    other.learn([{**ROW, "email": "omar.f@sample.org", "name": "Omar Fictional"}])
    gh.conflict_once = True  # alguien guardo antes: se vuelve a cargar y se unen los datos
    assert other.save()
    final = cm.ContactMemory(cm.GitHubStore("tok", session=gh), secret="tok")
    assert final.lookup("jane.testwell@example.com") and final.lookup("omar.f@sample.org")


def test_from_secrets_without_token_is_local(tmp_path):
    memory = cm.from_secrets({}, local_path=tmp_path / "m.bin")
    assert memory.label == "local"
