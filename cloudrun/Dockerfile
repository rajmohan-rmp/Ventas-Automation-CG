# Ventas Daily Dashboard — Cloud Run Job image
FROM python:3.12-slim

# LibreOffice (headless) renders the inline slide PNG; fonts for matplotlib/soffice.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY cloudrun/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# App code + brand assets
COPY ventas_slide_generator.py /app/
COPY run_daily.py              /app/
COPY assets/                   /app/assets/

# Writable scratch dir (Cloud Run filesystem is writable in-memory; /tmp is safe)
ENV VENTAS_WORKDIR=/tmp
# Recipients (override at deploy time with --set-env-vars)
ENV MAIL_TO="raj.mohan@6dtech.co.in"
ENV MAIL_CC="hari.krishnan@6dtech.co.in"
ENV MAIL_BCC=""

# Credentials are injected at deploy time via Secret Manager:
#   SA_KEY_JSON + GMAIL_DELEGATED_USER   (service account, domain-wide delegation)  [recommended]
#   or TOKEN_JSON                        (user OAuth token content)

ENTRYPOINT ["python", "run_daily.py"]
