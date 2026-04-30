import gzip
import hashlib
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from flask import Flask, request, redirect, url_for, send_file
from werkzeug.utils import secure_filename


APP_TITLE = "Обновление базы сварщиков"
SITE_DIR = Path(os.environ.get("WELDING_SITE_DIR", "/var/www/welding-site"))
UPLOAD_PASSWORD = os.environ.get("UPLOAD_PASSWORD", "smk_upload_1103")
MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "250"))

DB_FILENAME = "welding_db.json.gz"
VERSION_FILENAME = "db_version.json"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


def now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def target_db_path() -> Path:
    return SITE_DIR / DB_FILENAME


def target_version_path() -> Path:
    return SITE_DIR / VERSION_FILENAME


def backup_dir() -> Path:
    return SITE_DIR / "uploads_backup"


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_current_version() -> dict:
    p = target_version_path()
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def validate_database_gz(path: Path) -> tuple[dict, dict]:
    """
    Проверяет, что файл является gzip JSON-базой нашего сайта.
    Возвращает (db, version_info).
    """
    size_bytes = path.stat().st_size
    file_hash = sha256_file(path)

    try:
        with gzip.open(path, "rt", encoding="utf-8") as f:
            db = json.load(f)
    except Exception as e:
        raise ValueError(f"Файл не похож на корректный welding_db.json.gz: {e}")

    if not isinstance(db, dict):
        raise ValueError("Внутри файла должен быть JSON-объект.")

    required = ["meta", "welders", "joints"]
    missing = [key for key in required if key not in db]
    if missing:
        raise ValueError("В базе не хватает обязательных разделов: " + ", ".join(missing))

    if not isinstance(db.get("welders"), list) or not isinstance(db.get("joints"), list):
        raise ValueError("Поля welders и joints должны быть списками.")

    meta = db.get("meta") or {}
    generated_at = meta.get("generated_at") or now_iso()

    # version завязан на содержимое файла. Изменилась база — изменилась версия.
    version_info = {
        "version": file_hash,
        "generated_at": generated_at,
        "uploaded_at": now_iso(),
        "database_file": DB_FILENAME,
        "sha256": file_hash,
        "size_bytes": size_bytes,
        "welders": len(db.get("welders", [])),
        "joints": len(db.get("joints", [])),
        "period_start": meta.get("period_start"),
        "period_end": meta.get("period_end"),
    }
    return db, version_info


def make_backup() -> None:
    backup_dir().mkdir(parents=True, exist_ok=True)
    stamp = now_stamp()

    db_path = target_db_path()
    if db_path.exists():
        shutil.copy2(db_path, backup_dir() / f"welding_db_{stamp}.json.gz")

    ver_path = target_version_path()
    if ver_path.exists():
        shutil.copy2(ver_path, backup_dir() / f"db_version_{stamp}.json")


def atomic_replace(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    os.replace(src, dst)


def render_page(message: str = "", kind: str = "info") -> str:
    current = read_current_version()
    current_html = ""
    if current:
        current_html = f"""
        <div class="current">
          <div class="eyebrow">Текущая база на сервере</div>
          <div class="grid">
            <div><span>Загружено</span><b>{escape(current.get('uploaded_at') or '—')}</b></div>
            <div><span>Собрана</span><b>{escape(current.get('generated_at') or '—')}</b></div>
            <div><span>Сварщиков</span><b>{escape(current.get('welders') or '—')}</b></div>
            <div><span>Стыков</span><b>{escape(current.get('joints') or '—')}</b></div>
            <div><span>Период</span><b>{escape(current.get('period_start') or '—')} — {escape(current.get('period_end') or '—')}</b></div>
            <div><span>Размер</span><b>{format_size(current.get('size_bytes'))}</b></div>
          </div>
        </div>
        """
    else:
        current_html = """
        <div class="current">
          <div class="eyebrow">Текущая база на сервере</div>
          <p class="muted">Файл db_version.json пока не найден.</p>
        </div>
        """

    msg_html = f'<div class="msg {kind}">{escape(message)}</div>' if message else ""

    return f"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{APP_TITLE}</title>
  <style>
    :root {{
      --bg: #04080f;
      --panel: rgba(255,255,255,.055);
      --border: rgba(255,255,255,.10);
      --text: #f0f4ff;
      --muted: rgba(180,200,225,.62);
      --blue: #6ab3ff;
      --green: #34d399;
      --red: #ff6b8a;
      --yellow: #fbbf24;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      min-height: 100vh;
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      background:
        radial-gradient(circle at 15% 10%, rgba(82,160,255,.22), transparent 30%),
        radial-gradient(circle at 85% 8%, rgba(50,210,200,.12), transparent 28%),
        var(--bg);
      display: grid;
      place-items: center;
      padding: 24px;
    }}
    .card {{
      width: min(760px, 100%);
      border: 1px solid var(--border);
      border-radius: 28px;
      padding: 28px;
      background: var(--panel);
      backdrop-filter: blur(30px) saturate(160%);
      -webkit-backdrop-filter: blur(30px) saturate(160%);
      box-shadow: 0 24px 64px rgba(0,0,0,.48), inset 0 1px 0 rgba(255,255,255,.08);
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: clamp(24px, 4vw, 34px);
      letter-spacing: -.04em;
    }}
    .lead {{
      margin: 0 0 22px;
      color: var(--muted);
      line-height: 1.55;
    }}
    label {{
      display: block;
      margin: 16px 0 8px;
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: .1em;
      font-weight: 700;
    }}
    input[type=password], input[type=file] {{
      width: 100%;
      padding: 13px 14px;
      border-radius: 14px;
      color: var(--text);
      background: rgba(255,255,255,.06);
      border: 1px solid rgba(255,255,255,.12);
      outline: none;
    }}
    input[type=file] {{
      cursor: pointer;
    }}
    button, .link {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
      margin-top: 18px;
      padding: 0 18px;
      border-radius: 14px;
      border: 1px solid rgba(106,179,255,.35);
      color: white;
      background: linear-gradient(145deg, rgba(100,170,255,.90), rgba(58,120,230,.95));
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
    }}
    .link {{
      margin-left: 8px;
      background: rgba(255,255,255,.06);
      border-color: rgba(255,255,255,.12);
      color: var(--text);
    }}
    .msg {{
      margin: 18px 0;
      padding: 12px 14px;
      border-radius: 14px;
      border: 1px solid rgba(255,255,255,.12);
      background: rgba(255,255,255,.055);
      line-height: 1.45;
    }}
    .msg.success {{ color: var(--green); border-color: rgba(52,211,153,.28); background: rgba(52,211,153,.08); }}
    .msg.error {{ color: var(--red); border-color: rgba(255,107,138,.28); background: rgba(255,107,138,.08); }}
    .msg.info {{ color: var(--yellow); border-color: rgba(251,191,36,.28); background: rgba(251,191,36,.08); }}
    .current {{
      margin-top: 20px;
      padding: 16px;
      border: 1px solid rgba(255,255,255,.10);
      border-radius: 18px;
      background: rgba(255,255,255,.035);
    }}
    .eyebrow {{
      margin-bottom: 12px;
      color: var(--muted);
      font-size: 12px;
      letter-spacing: .14em;
      text-transform: uppercase;
      font-weight: 800;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 10px;
    }}
    .grid div {{
      padding: 10px;
      border-radius: 12px;
      background: rgba(255,255,255,.045);
    }}
    .grid span {{
      display: block;
      color: var(--muted);
      font-size: 12px;
    }}
    .grid b {{
      display: block;
      margin-top: 3px;
      word-break: break-word;
    }}
    .muted {{ color: var(--muted); }}
    @media (max-width: 640px) {{
      body {{ padding: 14px; }}
      .card {{ padding: 20px; border-radius: 22px; }}
      .grid {{ grid-template-columns: 1fr; }}
      .link {{ margin-left: 0; width: 100%; }}
      button {{ width: 100%; }}
    }}
  </style>
</head>
<body>
  <main class="card">
    <h1>Обновление базы</h1>
    <p class="lead">Загрузите файл <b>welding_db.json.gz</b>. После успешной загрузки база обновится на сервере для всех пользователей сайта.</p>
    {msg_html}

    <form method="post" enctype="multipart/form-data">
      <label for="password">Пароль загрузки</label>
      <input id="password" name="password" type="password" autocomplete="current-password" required>

      <label for="database">Файл базы</label>
      <input id="database" name="database" type="file" accept=".gz,application/gzip" required>

      <button type="submit">Обновить базу</button>
      <a class="link" href="/">Вернуться на сайт</a>
    </form>

    {current_html}
  </main>
</body>
</html>"""


def escape(value) -> str:
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def format_size(value) -> str:
    try:
        size = int(value)
    except Exception:
        return "—"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} КБ"
    return f"{size / 1024 / 1024:.1f} МБ"


@app.get("/admin-upload")
@app.get("/admin-upload/")
def upload_form():
    return render_page()


@app.post("/admin-upload")
@app.post("/admin-upload/")
def upload_database():
    password = request.form.get("password", "")
    if password != UPLOAD_PASSWORD:
        return render_page("Неверный пароль загрузки.", "error"), 403

    file = request.files.get("database")
    if not file or not file.filename:
        return render_page("Файл не выбран.", "error"), 400

    filename = secure_filename(file.filename)
    if not filename.endswith(".json.gz") and not filename.endswith(".gz"):
        return render_page("Нужен файл welding_db.json.gz.", "error"), 400

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".json.gz") as tmp:
            tmp_path = Path(tmp.name)
            file.save(tmp)

        db, version_info = validate_database_gz(tmp_path)

        # Если файл валиден — делаем резервную копию текущей базы и заменяем.
        make_backup()

        target_tmp = SITE_DIR / f".{DB_FILENAME}.uploading"
        shutil.copy2(tmp_path, target_tmp)
        atomic_replace(target_tmp, target_db_path())

        version_tmp = SITE_DIR / f".{VERSION_FILENAME}.uploading"
        version_tmp.write_text(json.dumps(version_info, ensure_ascii=False, indent=2), encoding="utf-8")
        atomic_replace(version_tmp, target_version_path())

        message = (
            f"Готово. База обновлена. "
            f"Сварщиков: {len(db.get('welders', []))}. "
            f"Стыков: {len(db.get('joints', []))}. "
            f"Версия: {version_info['version'][:16]}…"
        )
        return render_page(message, "success")

    except Exception as e:
        return render_page(f"Ошибка загрузки: {e}", "error"), 400

    finally:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass


@app.get("/admin-upload/health")
def health():
    return {"ok": True, "site_dir": str(SITE_DIR)}


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8080)
