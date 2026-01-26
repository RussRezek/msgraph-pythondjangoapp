#!/bin/bash
set -e

# Collect static files
python /usr/src/app/manage.py collectstatic --noinput 2>/dev/null || true

# Debugging is handled inside the WSGI process (see graph_main_app/wsgi.py)

# Start gunicorn via python so we can pass the interpreter flag to disable frozen modules
# This prevents debugpy from warning about frozen modules and helps breakpoints work reliably
exec python -Xfrozen_modules=off -m gunicorn graph_main_app.wsgi:application \
    --bind 0.0.0.0:8000 \
    --workers $GUNICORN_WORKERS \
    --threads $GUNICORN_THREADS \
    --timeout $GUNICORN_TIMEOUT
