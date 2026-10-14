#!/usr/bin/env bash
# One-shot deploy of the Ventas daily dashboard to Cloud Run Jobs + Cloud Scheduler.
# Run from the PROJECT ROOT (the folder that contains run_daily.py), on a machine
# with gcloud installed and authenticated:  gcloud auth login
set -euo pipefail

# ---- EDIT THESE ----
PROJECT_ID="your-gcp-project-id"
REGION="asia-south1"                       # Mumbai; pick your nearest region
DELEGATED_USER="raj.mohan@6dtech.co.in"    # mailbox the job reads & sends as
MAIL_TO="raj.mohan@6dtech.co.in"
MAIL_CC="hari.krishnan@6dtech.co.in"
MAIL_BCC=""
SCHEDULE="0 10 * * *"                       # 10:00 daily
TZ="Asia/Kolkata"
# --------------------

JOB="ventas-daily-dashboard"
SA="ventas-job"
SA_EMAIL="${SA}@${PROJECT_ID}.iam.gserviceaccount.com"
IMAGE="${REGION}-docker.pkg.dev/${PROJECT_ID}/ventas/ventas-job:latest"

gcloud config set project "$PROJECT_ID"

echo "==> Enable APIs"
gcloud services enable run.googleapis.com cloudscheduler.googleapis.com \
  artifactregistry.googleapis.com secretmanager.googleapis.com \
  gmail.googleapis.com drive.googleapis.com cloudbuild.googleapis.com

echo "==> Service account (the job's identity)"
gcloud iam service-accounts create "$SA" \
  --display-name="Ventas daily job" 2>/dev/null || true

echo "==> Artifact Registry repo"
gcloud artifacts repositories create ventas --repository-format=docker \
  --location="$REGION" 2>/dev/null || true

echo "==> Build & push image (uses cloudrun/Dockerfile)"
gcloud builds submit --config=- . <<'YAML'
steps:
- name: gcr.io/cloud-builders/docker
  args: ['build','-f','cloudrun/Dockerfile','-t','$_IMAGE','.']
- name: gcr.io/cloud-builders/docker
  args: ['push','$_IMAGE']
substitutions:
  _IMAGE: '${IMAGE}'
images: ['${IMAGE}']
YAML

echo "==> Store the service-account key for domain-wide delegation in Secret Manager"
echo "    (create the key + enable DWD first - see DEPLOY_CLOUDRUN.md). Then:"
echo "    gcloud secrets create ventas-sa-key --data-file=sa-key.json"
gcloud secrets add-iam-policy-binding ventas-sa-key \
  --member="serviceAccount:${SA_EMAIL}" \
  --role="roles/secretmanager.secretAccessor" 2>/dev/null || true

echo "==> Create / update the Cloud Run Job"
gcloud run jobs deploy "$JOB" \
  --image="$IMAGE" --region="$REGION" \
  --service-account="$SA_EMAIL" \
  --max-retries=1 --task-timeout=600s --memory=1Gi \
  --set-env-vars="GMAIL_DELEGATED_USER=${DELEGATED_USER},MAIL_TO=${MAIL_TO},MAIL_CC=${MAIL_CC},MAIL_BCC=${MAIL_BCC},VENTAS_WORKDIR=/tmp" \
  --set-secrets="SA_KEY_JSON=ventas-sa-key:latest"

echo "==> Allow the service account to execute the job (for Scheduler)"
gcloud run jobs add-iam-policy-binding "$JOB" --region="$REGION" \
  --member="serviceAccount:${SA_EMAIL}" --role="roles/run.invoker" 2>/dev/null || true

echo "==> Schedule it daily via Cloud Scheduler (invokes the job)"
gcloud scheduler jobs create http "${JOB}-trigger" \
  --location="$REGION" --schedule="$SCHEDULE" --time-zone="$TZ" \
  --uri="https://${REGION}-run.googleapis.com/apis/run.googleapis.com/v1/namespaces/${PROJECT_ID}/jobs/${JOB}:run" \
  --http-method=POST \
  --oauth-service-account-email="$SA_EMAIL" 2>/dev/null \
  || gcloud scheduler jobs update http "${JOB}-trigger" \
       --location="$REGION" --schedule="$SCHEDULE" --time-zone="$TZ"

echo "==> Done. Test now with:  gcloud run jobs execute ${JOB} --region ${REGION}"
