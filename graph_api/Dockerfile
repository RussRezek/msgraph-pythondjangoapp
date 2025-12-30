# Debian 12  (Bookworm)
# Production Dockerfile for Django (multi-stage optional kept simple)
FROM python:3.12-slim AS runtime

# Ensure deterministic, secure Python behavior
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PIP_NO_CACHE_DIR=1

# allow runtime configuration overrides
ENV PIP_DISABLE_PIP_VERSION_CHECK=1

ENV GUNICORN_WORKERS=2
ENV GUNICORN_THREADS=2
ENV GUNICORN_TIMEOUT=60

# Install system deps required for building common Python packages and postgres client libs
RUN apt-get update && apt-get install --no-install-recommends -y build-essential libpq-dev curl gosu ca-certificates wget gnupg2 apt-transport-https emacs && rm -rf /var/lib/apt/lists/*


# Import Microsoft signing key into a keyring file (recommended modern approach)
RUN curl -fsSL https://packages.microsoft.com/keys/microsoft.asc \
    | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg

# Add Microsoft repo (Debian 12 Bookworm EXAMPLE; see notes to switch to Debian 13/Trixie)
RUN curl -fsSL https://packages.microsoft.com/config/debian/12/prod.list \
    | tee /etc/apt/sources.list.d/mssql-release.list

# Install driver & tools (accept EULA)
RUN apt-get update && \
    ACCEPT_EULA=Y apt-get install --no-install-recommends -y \
      msodbcsql18 mssql-tools18 unixodbc-dev && \
    # put tools on PATH for all shells
    echo 'export PATH="/opt/mssql-tools18/bin:$PATH"' > /etc/profile.d/mssql-tools.sh && \
    # Accept EULA automatically for future driver updates (per MS docs)
    mkdir -p /opt/microsoft/msodbcsql18 && \
    touch /opt/microsoft/msodbcsql18/ACCEPT_EULA && \
    rm -rf /var/lib/apt/lists/*


# Create a non-root user
RUN groupadd --system --gid 3001 appuser && useradd  --system --uid 3001 --gid 3001 --create-home --home-dir /home/appuser --shell /usr/sbin/nologin appuser

# Install virtalenv
RUN pip install virtualenv

ENV VIRTUAL_ENV=/opt/virtualenv
RUN python -m virtualenv $VIRTUAL_ENV
ENV PATH="$VIRTUAL_ENV/bin:$PATH"

# Install dependencies:
COPY requirements.txt .
RUN pip install -r requirements.txt

# Create workdir
WORKDIR /usr/src/app

# Copy project source
COPY . /usr/src/app
COPY secrets/. /usr/src/app

# Prepare directories for static/media and ensure permissions for non-root user
RUN mkdir -p /usr/src/app/staticfiles /usr/src/app/media && chown -R appuser:appuser /usr/src/app

# Expose application port (documentation)
EXPOSE 8000

# Default command - runs as root, collects statics, then switches to appuser
CMD ["sh", "-c", "python /usr/src/app/manage.py collectstatic --noinput 2>/dev/null || true && chown -R appuser:appuser /usr/src/app/staticfiles /usr/src/app/media && gosu appuser gunicorn graph_main_app.wsgi:application --bind 0.0.0.0:8000"]
