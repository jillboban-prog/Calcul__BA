"""calculBA — API et service de l'interface.

Lancer depuis la racine du projet :
    uvicorn backend.main:app --reload
"""
from __future__ import annotations

import json
import logging
import re
import sqlite3
from contextlib import asynccontextmanager
from datetime import date
from pathlib import Path
from typing import Any, Literal, Optional
from urllib.parse import quote

from fastapi import Depends, FastAPI, File, Form, HTTPException, Query, Request, Response, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field, field_validator
from starlette.exceptions import HTTPException as StarletteHTTPException

from . import analysis, settings
from .db import get_db, init_db, ts
from .security import (
    DUMMY_HASH, close_session, current_user, hash_password, login_throttle,
    needs_rehash, open_session, public_user, verify_password,
)

log = logging.getLogger("calculba")

if not settings.STATIC_DIR.is_dir():
    raise RuntimeError(
        f"Dossier introuvable : {settings.STATIC_DIR}\n"
        "Lancez calculBA depuis la racine du projet, qui doit contenir backend/ et frontend/."
    )


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    yield


API_DESCRIPTION = """
API de calculBA. Toutes les routes sauf **Comptes** et **Service** exigent d'être connecté.

**Pour essayer les routes ici :** connectez-vous d'abord dans un autre onglet sur
[la page d'accueil](/) ; le cookie de session est partagé avec cette page.
"""

TAGS = [
    {"name": "Comptes", "description": "Inscription, connexion, session et suppression du compte."},
    {"name": "Jeux de données", "description": "Créer, importer, renommer, supprimer un jeu."},
    {"name": "Observations", "description": "Lignes d'un jeu : lister (paginé), ajouter, modifier, supprimer."},
    {"name": "Analyse", "description": "Statistiques descriptives, données des graphiques et exports."},
    {"name": "Service", "description": "État du serveur et paramètres de l'interface."},
]

# docs_url=None : la page /docs par défaut charge Swagger depuis Internet, ce que la
# politique de sécurité (CSP) bloque. On sert une version locale plus bas.
app = FastAPI(title="calculBA — données de génie civil", version=settings.VERSION, lifespan=lifespan,
              description=API_DESCRIPTION, openapi_tags=TAGS, docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=settings.STATIC_DIR), name="static")

# --------------------------------------------------------------- middlewares

CSP = (
    "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
    "img-src 'self' data: blob:; font-src 'self'; connect-src 'self'; "
    "frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
)


@app.middleware("http")
async def protect(request: Request, call_next):
    # Protection CSRF : tout appel qui modifie des données doit porter un en-tête
    # personnalisé, qu'un formulaire d'un autre site ne peut pas ajouter.
    if (request.url.path.startswith("/api/") and request.method in {"POST", "PUT", "PATCH", "DELETE"}
            and request.headers.get(settings.CSRF_HEADER) != settings.CSRF_VALUE):
        return JSONResponse({"detail": "Requête refusée (en-tête de sécurité manquant)."}, status_code=403)
    response = await call_next(request)
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "DENY")
    response.headers.setdefault("Referrer-Policy", "same-origin")
    response.headers.setdefault("Content-Security-Policy", CSP)
    if request.url.path.startswith("/api/"):
        response.headers.setdefault("Cache-Control", "no-store")
    return response


# --------------------------------------------- erreurs toujours au format JSON

FIELD_NAMES = {
    "full_name": "nom complet", "email": "adresse e-mail", "password": "mot de passe",
    "birth_date": "date de naissance", "name": "nom du jeu", "domain": "domaine",
    "columns": "colonnes", "values": "valeurs", "current_password": "mot de passe actuel",
}


@app.exception_handler(RequestValidationError)
async def validation_error(_: Request, exc: RequestValidationError):
    err = exc.errors()[0] if exc.errors() else {}
    loc = [str(p) for p in err.get("loc", []) if p not in ("body", "query", "path")]
    field = FIELD_NAMES.get(loc[0], loc[0]) if loc else None
    msg = err.get("msg", "valeur invalide").removeprefix("Value error, ")
    text = f"Champ « {field} » : {msg}" if field else f"Requête invalide : {msg}"
    return JSONResponse({"detail": text}, status_code=422)


@app.exception_handler(StarletteHTTPException)
async def http_error(_: Request, exc: StarletteHTTPException):
    detail = exc.detail if isinstance(exc.detail, str) else "Erreur."
    if exc.status_code == 404 and detail == "Not Found":
        detail = "Ressource introuvable."
    return JSONResponse({"detail": detail}, status_code=exc.status_code, headers=getattr(exc, "headers", None))


@app.exception_handler(Exception)
async def unexpected_error(_: Request, exc: Exception):
    log.exception("Erreur interne", exc_info=exc)
    return JSONResponse({"detail": "Erreur interne du serveur. Consultez le terminal pour le détail."}, status_code=500)


# ------------------------------------------------------------------- modèles

EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]{2,}$")


class RegisterIn(BaseModel):
    full_name: str = Field(min_length=2, max_length=120)
    email: str = Field(max_length=254)
    password: str = Field(min_length=settings.PASSWORD_MIN, max_length=settings.PASSWORD_MAX)
    birth_date: Optional[str] = None

    @field_validator("full_name")
    @classmethod
    def _name(cls, v: str) -> str:
        v = " ".join(v.split())
        if len(v) < 2:
            raise ValueError("saisissez votre nom complet.")
        return v

    @field_validator("email")
    @classmethod
    def _email(cls, v: str) -> str:
        v = v.strip().lower()
        if not EMAIL_RE.match(v):
            raise ValueError("adresse invalide.")
        return v

    @field_validator("birth_date")
    @classmethod
    def _birth(cls, v: Optional[str]) -> Optional[str]:
        if v is None or not v.strip():
            return None
        try:
            d = date.fromisoformat(v.strip())
        except ValueError:
            raise ValueError("date invalide.")
        today = date.today()
        if not (date(today.year - 120, 1, 1) <= d < today):
            raise ValueError("date hors de l'intervalle plausible.")
        return d.isoformat()


class LoginIn(BaseModel):
    email: str = Field(max_length=254)
    password: str = Field(max_length=settings.PASSWORD_MAX)


class PasswordIn(BaseModel):
    current_password: str = Field(max_length=settings.PASSWORD_MAX)


Domain = Literal[tuple(settings.DOMAINS)]  # type: ignore[valid-type]


class DatasetIn(BaseModel):
    name: str = Field(min_length=1, max_length=settings.MAX_NAME_LENGTH)
    domain: Domain = "Autre"
    columns: list[str] = Field(min_length=1, max_length=settings.MAX_COLUMNS)

    @field_validator("name")
    @classmethod
    def _name(cls, v: str) -> str:
        v = v.strip()
        if not v:
            raise ValueError("donnez un nom au jeu de données.")
        return v

    @field_validator("columns")
    @classmethod
    def _cols(cls, v: list[str]) -> list[str]:
        try:
            return analysis.validate_columns(v)
        except analysis.ImportProblem as e:
            raise ValueError(str(e))


class DatasetPatch(BaseModel):
    name: Optional[str] = Field(default=None, min_length=1, max_length=settings.MAX_NAME_LENGTH)
    domain: Optional[Domain] = None


Scalar = Optional[str | int | float]


class RecordIn(BaseModel):
    values: dict[str, Scalar]


# ------------------------------------------------------------------- helpers

def load_dataset(conn: sqlite3.Connection, did: int, user_id: int) -> dict:
    row = conn.execute(
        """SELECT d.*, (SELECT COUNT(*) FROM records r WHERE r.dataset_id = d.id) AS record_count
           FROM datasets d WHERE d.id = ? AND d.user_id = ?""",
        (did, user_id),
    ).fetchone()
    if not row:
        raise HTTPException(404, "Ce jeu de données n'existe pas ou appartient à un autre compte.")
    return dataset_out(row)


def dataset_out(row: sqlite3.Row) -> dict:
    return {
        "id": row["id"], "name": row["name"], "domain": row["domain"],
        "columns": json.loads(row["columns_json"]),
        "record_count": row["record_count"],
        "created_at": row["created_at"], "updated_at": row["updated_at"] or row["created_at"],
    }


def load_rows(conn: sqlite3.Connection, did: int) -> list[dict]:
    return [json.loads(r[0]) for r in conn.execute("SELECT values_json FROM records WHERE dataset_id = ? ORDER BY id", (did,))]


def clean_record(columns: list[str], values: dict[str, Any]) -> dict:
    unknown = set(values) - set(columns)
    if unknown:
        raise HTTPException(400, f"Colonne inconnue : {', '.join(sorted(unknown))}.")
    clean = {c: analysis.clean_value(values.get(c)) for c in columns}
    if all(v is None for v in clean.values()):
        raise HTTPException(400, "L'observation est vide : renseignez au moins une valeur.")
    return clean


def check_numeric(conn: sqlite3.Connection, did: int, columns: list[str], clean: dict, skip_rid: int | None = None) -> None:
    """Refuse un texte dans une colonne détectée comme numérique (ex. « abc » dans Resistance_MPa).

    Seules les colonnes où la nouvelle valeur n'est pas un nombre sont examinées,
    pour rester rapide sur les gros jeux. Une colonne encore vide accepte tout.
    """
    suspects = [c for c in columns if clean.get(c) is not None and analysis.to_number(clean[c]) is None]
    if not suspects:
        return
    rows = [json.loads(r["values_json"]) for r in conn.execute(
        "SELECT id, values_json FROM records WHERE dataset_id = ? AND id IS NOT ?", (did, skip_rid))]
    prof = analysis.profile(suspects, rows)
    bad = [c for c in suspects if prof[c]["numeric"]]
    if bad:
        raise HTTPException(400, f"Valeur non numérique pour : {', '.join(bad)}. Corrigez-la (ex. 30,5) ou laissez le champ vide.")


def touch(conn: sqlite3.Connection, did: int) -> None:
    conn.execute("UPDATE datasets SET updated_at = ? WHERE id = ?", (ts(), did))


def require_column(ds: dict, col: str) -> str:
    if col not in ds["columns"]:
        raise HTTPException(400, f"La colonne « {col} » n'existe pas dans ce jeu.")
    return col


# --------------------------------------------------------------------- pages

@app.get("/", include_in_schema=False)
def index():
    return FileResponse(settings.FRONTEND_DIR / "index.html", headers={"Cache-Control": "no-cache"})


@app.get("/docs", include_in_schema=False)
def api_docs():
    """Documentation interactive de l'API (Swagger UI hébergé localement)."""
    return FileResponse(settings.FRONTEND_DIR / "api-docs.html", headers={"Cache-Control": "no-cache"})


@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return FileResponse(settings.STATIC_DIR / "favicon.svg", media_type="image/svg+xml")


@app.get("/api/health", tags=["Service"], summary="Vérifier que le serveur répond")
def health():
    return {"status": "ok", "application": settings.APP_NAME, "version": settings.VERSION}


@app.get("/api/meta", tags=["Service"], summary="Domaines et limites (pour l'interface)")
def meta():
    return {
        "version": settings.VERSION, "domains": settings.DOMAINS,
        "limits": {"rows": settings.MAX_ROWS, "columns": settings.MAX_COLUMNS,
                   "upload_mb": settings.MAX_UPLOAD_BYTES // (1024 * 1024),
                   "password_min": settings.PASSWORD_MIN},
    }


# ---------------------------------------------------------------------- auth

@app.post("/api/auth/register", status_code=201, tags=["Comptes"], summary="Créer un compte (ouvre une session)")
def register(body: RegisterIn, response: Response, conn: sqlite3.Connection = Depends(get_db)):
    try:
        cur = conn.execute(
            "INSERT INTO users (full_name, birth_date, email, password_hash, created_at) VALUES (?, ?, ?, ?, ?)",
            (body.full_name, body.birth_date, body.email, hash_password(body.password), ts()),
        )
    except sqlite3.IntegrityError:
        raise HTTPException(409, "Un compte existe déjà avec cette adresse e-mail. Connectez-vous.")
    open_session(conn, response, cur.lastrowid)
    return {"user": {"id": cur.lastrowid, "full_name": body.full_name, "email": body.email}}


@app.post("/api/auth/login", tags=["Comptes"], summary="Se connecter")
def login(body: LoginIn, request: Request, response: Response, conn: sqlite3.Connection = Depends(get_db)):
    email = body.email.strip().lower()
    key = f"{request.client.host if request.client else '?'}|{email}"
    login_throttle.check(key)
    row = conn.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()
    ok = verify_password(body.password, row["password_hash"] if row else DUMMY_HASH)
    if not row or not ok:
        login_throttle.fail(key)
        raise HTTPException(401, "Adresse e-mail ou mot de passe incorrect.")
    login_throttle.reset(key)
    if needs_rehash(row["password_hash"]):
        conn.execute("UPDATE users SET password_hash = ? WHERE id = ?", (hash_password(body.password), row["id"]))
    open_session(conn, response, row["id"])
    return {"user": public_user(row)}


@app.post("/api/auth/logout", tags=["Comptes"], summary="Se déconnecter")
def logout(request: Request, response: Response, conn: sqlite3.Connection = Depends(get_db)):
    close_session(conn, request, response)
    return {"message": "Vous êtes déconnecté."}


@app.get("/api/auth/me", tags=["Comptes"], summary="Utilisateur connecté")
def me(user: dict = Depends(current_user)):
    return {"user": user}


@app.post("/api/auth/delete-account", tags=["Comptes"], summary="Supprimer son compte et ses données")
def delete_account(body: PasswordIn, request: Request, response: Response,
                   user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    row = conn.execute("SELECT password_hash FROM users WHERE id = ?", (user["id"],)).fetchone()
    if not verify_password(body.current_password, row["password_hash"]):
        raise HTTPException(403, "Mot de passe incorrect.")
    ids = [r[0] for r in conn.execute("SELECT id FROM datasets WHERE user_id = ?", (user["id"],))]
    conn.executemany("DELETE FROM records WHERE dataset_id = ?", [(i,) for i in ids])
    conn.execute("DELETE FROM datasets WHERE user_id = ?", (user["id"],))
    conn.execute("DELETE FROM user_sessions WHERE user_id = ?", (user["id"],))
    conn.execute("DELETE FROM users WHERE id = ?", (user["id"],))
    conn.commit()
    close_session(conn, request, response)
    return {"message": "Compte et données supprimés."}


# ------------------------------------------------------------------ datasets

@app.get("/api/datasets", tags=["Jeux de données"], summary="Lister mes jeux")
def list_datasets(user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    rows = conn.execute(
        """SELECT d.*, (SELECT COUNT(*) FROM records r WHERE r.dataset_id = d.id) AS record_count
           FROM datasets d WHERE d.user_id = ? ORDER BY COALESCE(d.updated_at, d.created_at) DESC, d.id DESC""",
        (user["id"],),
    ).fetchall()
    return [dataset_out(r) for r in rows]


@app.post("/api/datasets", status_code=201, tags=["Jeux de données"], summary="Créer un jeu vide (colonnes)")
def create_dataset(body: DatasetIn, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    now = ts()
    cur = conn.execute(
        "INSERT INTO datasets (user_id, name, domain, columns_json, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?)",
        (user["id"], body.name, body.domain, json.dumps(body.columns, ensure_ascii=False), now, now),
    )
    conn.commit()
    return load_dataset(conn, cur.lastrowid, user["id"])


@app.post("/api/datasets/import", status_code=201, tags=["Jeux de données"], summary="Importer un fichier CSV ou Excel")
async def import_dataset(
    file: UploadFile = File(...),
    name: str = Form(""),
    domain: str = Form("Autre"),
    user: dict = Depends(current_user),
    conn: sqlite3.Connection = Depends(get_db),
):
    raw = await file.read(settings.MAX_UPLOAD_BYTES + 1)
    if len(raw) > settings.MAX_UPLOAD_BYTES:
        raise HTTPException(413, f"Fichier trop volumineux (maximum {settings.MAX_UPLOAD_BYTES // (1024 * 1024)} Mo).")
    filename = file.filename or "donnees.csv"
    try:
        columns, data = analysis.read_table(filename, raw)
    except analysis.ImportProblem as e:
        raise HTTPException(400, str(e))
    if domain not in settings.DOMAINS:
        domain = "Autre"
    ds_name = (name.strip() or Path(filename).stem or "Données importées")[: settings.MAX_NAME_LENGTH]
    now = ts()
    cur = conn.execute(
        "INSERT INTO datasets (user_id, name, domain, columns_json, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?)",
        (user["id"], ds_name, domain, json.dumps(columns, ensure_ascii=False), now, now),
    )
    did = cur.lastrowid
    conn.executemany(
        "INSERT INTO records (dataset_id, values_json, created_at) VALUES (?, ?, ?)",
        ((did, json.dumps(dict(zip(columns, row)), ensure_ascii=False), now) for row in data),
    )
    conn.commit()
    ds = load_dataset(conn, did, user["id"])
    return {"dataset": ds, "rows_imported": len(data)}


@app.get("/api/datasets/{did}", tags=["Jeux de données"], summary="Détail d'un jeu")
def get_dataset(did: int, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    return load_dataset(conn, did, user["id"])


@app.patch("/api/datasets/{did}", tags=["Jeux de données"], summary="Renommer / changer le domaine")
def update_dataset(did: int, body: DatasetPatch, user: dict = Depends(current_user),
                   conn: sqlite3.Connection = Depends(get_db)):
    load_dataset(conn, did, user["id"])
    if body.name is not None:
        if not body.name.strip():
            raise HTTPException(400, "Donnez un nom au jeu de données.")
        conn.execute("UPDATE datasets SET name = ? WHERE id = ?", (body.name.strip(), did))
    if body.domain is not None:
        conn.execute("UPDATE datasets SET domain = ? WHERE id = ?", (body.domain, did))
    touch(conn, did)
    conn.commit()
    return load_dataset(conn, did, user["id"])


@app.delete("/api/datasets/{did}", tags=["Jeux de données"], summary="Supprimer un jeu")
def delete_dataset(did: int, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    load_dataset(conn, did, user["id"])
    conn.execute("DELETE FROM records WHERE dataset_id = ?", (did,))
    conn.execute("DELETE FROM datasets WHERE id = ? AND user_id = ?", (did, user["id"]))
    conn.commit()
    return {"message": "Jeu de données supprimé."}


# ------------------------------------------------------------------- records

@app.get("/api/datasets/{did}/records", tags=["Observations"], summary="Lister les observations (paginé)")
def list_records(did: int, offset: int = Query(0, ge=0), limit: int = Query(50, ge=1, le=500),
                 user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    rows = conn.execute(
        "SELECT id, values_json, created_at FROM records WHERE dataset_id = ? ORDER BY id LIMIT ? OFFSET ?",
        (did, limit, offset),
    ).fetchall()
    # Les valeurs sont imbriquées dans « values » : une colonne nommée « id »
    # ne peut plus écraser l'identifiant de l'observation.
    return {
        "total": ds["record_count"], "offset": offset, "limit": limit,
        "records": [{"id": r["id"], "created_at": r["created_at"], "values": json.loads(r["values_json"])} for r in rows],
    }


@app.post("/api/datasets/{did}/records", status_code=201, tags=["Observations"], summary="Ajouter une observation")
def add_record(did: int, body: RecordIn, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    if ds["record_count"] >= settings.MAX_ROWS:
        raise HTTPException(400, f"Ce jeu a atteint la limite de {settings.MAX_ROWS} observations.")
    clean = clean_record(ds["columns"], body.values)
    check_numeric(conn, did, ds["columns"], clean)
    cur = conn.execute("INSERT INTO records (dataset_id, values_json, created_at) VALUES (?, ?, ?)",
                       (did, json.dumps(clean, ensure_ascii=False), ts()))
    touch(conn, did)
    conn.commit()
    return {"id": cur.lastrowid, "values": clean}


@app.put("/api/datasets/{did}/records/{rid}", tags=["Observations"], summary="Modifier une observation")
def update_record(did: int, rid: int, body: RecordIn, user: dict = Depends(current_user),
                  conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    clean = clean_record(ds["columns"], body.values)
    check_numeric(conn, did, ds["columns"], clean, skip_rid=rid)
    cur = conn.execute("UPDATE records SET values_json = ? WHERE id = ? AND dataset_id = ?",
                       (json.dumps(clean, ensure_ascii=False), rid, did))
    if cur.rowcount == 0:
        raise HTTPException(404, "Observation introuvable.")
    touch(conn, did)
    conn.commit()
    return {"id": rid, "values": clean}


@app.delete("/api/datasets/{did}/records/{rid}", tags=["Observations"], summary="Supprimer une observation")
def delete_record(did: int, rid: int, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    load_dataset(conn, did, user["id"])
    cur = conn.execute("DELETE FROM records WHERE id = ? AND dataset_id = ?", (rid, did))
    if cur.rowcount == 0:
        raise HTTPException(404, "Observation introuvable.")
    touch(conn, did)
    conn.commit()
    return {"message": "Observation supprimée."}


# ------------------------------------------------------------------ analysis

@app.get("/api/datasets/{did}/stats", tags=["Analyse"], summary="Statistiques descriptives")
def stats(did: int, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    return {"dataset": ds, **analysis.describe(ds["columns"], load_rows(conn, did))}


@app.get("/api/datasets/{did}/charts/series", tags=["Analyse"], summary="Données : valeurs dans l'ordre de saisie")
def chart_series(did: int, column: str, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    return analysis.series(load_rows(conn, did), require_column(ds, column))


@app.get("/api/datasets/{did}/charts/scatter", tags=["Analyse"], summary="Données : nuage X/Y et régression")
def chart_scatter(did: int, x: str, y: str, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    return analysis.scatter(load_rows(conn, did), require_column(ds, x), require_column(ds, y))


@app.get("/api/datasets/{did}/charts/groups", tags=["Analyse"], summary="Données : moyenne par groupe")
def chart_groups(did: int, by: str, value: str, user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    return analysis.groups(load_rows(conn, did), require_column(ds, by), require_column(ds, value))


@app.get("/api/datasets/{did}/export", tags=["Analyse"], summary="Exporter (xlsx, csv, csv-excel)")
def export(did: int, format: Literal["xlsx", "csv", "csv-excel"] = "xlsx",
           user: dict = Depends(current_user), conn: sqlite3.Connection = Depends(get_db)):
    ds = load_dataset(conn, did, user["id"])
    rows = load_rows(conn, did)
    numeric = analysis.profile(ds["columns"], rows)
    numeric_cols = [c for c, p in numeric.items() if p["numeric"]]
    table = analysis.export_rows(ds["columns"], rows, numeric_cols)
    base = re.sub(r"[^\w\-]+", "_", ds["name"], flags=re.UNICODE).strip("_") or f"jeu_{did}"
    if format == "xlsx":
        content = analysis.to_xlsx(ds["name"], ds["columns"], table)
        media, ext = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"
    else:
        content = analysis.to_csv(ds["columns"], table, excel_fr=(format == "csv-excel"))
        media, ext = "text/csv; charset=utf-8", "csv"
    filename = f"calculBA_{base}.{ext}"
    ascii_name = filename.encode("ascii", "ignore").decode() or f"calculBA_{did}.{ext}"
    return Response(content, media_type=media, headers={
        "Content-Disposition": f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(filename)}"})
