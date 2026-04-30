# Этап 1 — статический сайт сварщиков

В этой папке готовый комплект для первого запуска сайта.

## Что внутри

- `index.html` — сайт.
- `welding_db.json.gz` — база, которую сайт автоматически подхватывает.
- `build_welding_db.py` — сборщик базы из Excel.
- `build_welding_db.cmd` — запуск сборщика на Windows из сетевой папки.
- `nginx_welding_site.conf` — пример конфига Nginx для Ubuntu.

## Как проверить локально

Открой PowerShell в этой папке и запусти:

```powershell
py -3 -m http.server 8000
```

Потом открой:

```text
http://localhost:8000
```

На боевой версии кнопка «Подгрузить базу» на главной скрыта. Сайт должен сам загрузить `welding_db.json.gz`.

## Как обновлять базу локально

Положи `build_welding_db.py` и `build_welding_db.cmd` рядом с файлом `Статистика*.xlsm`, затем запусти `build_welding_db.cmd`.

Он создаст:

```text
welding_db.json
welding_db.json.gz
```

Для сайта нужен `welding_db.json.gz`.

## Минимальные команды на сервере Ubuntu

Если на сервере уже есть Nginx и Git:

```bash
cd /var/www
sudo git clone https://github.com/USER/REPO.git welding-site
sudo cp /var/www/welding-site/nginx_welding_site.conf /etc/nginx/sites-available/welding-site
sudo ln -sf /etc/nginx/sites-available/welding-site /etc/nginx/sites-enabled/welding-site
sudo nginx -t
sudo systemctl reload nginx
```

Если Nginx/Git ещё не стоят:

```bash
sudo apt update
sudo apt install -y nginx git
```

Заменить `USER/REPO` на адрес твоего репозитория GitHub.
