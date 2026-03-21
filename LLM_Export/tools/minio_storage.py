import os
import logging
import threading
import time

log = logging.getLogger("minio_storage")

STORAGE_BACKEND = os.getenv("STORAGE_BACKEND", "local").strip().lower()

MINIO_ENDPOINT = os.getenv("MINIO_ENDPOINT", "")
MINIO_PUBLIC_ENDPOINT = os.getenv("MINIO_PUBLIC_ENDPOINT", "")
MINIO_ACCESS_KEY = os.getenv("MINIO_ACCESS_KEY", "")
MINIO_SECRET_KEY = os.getenv("MINIO_SECRET_KEY", "")
MINIO_BUCKET = os.getenv("MINIO_BUCKET", "file-exports")
MINIO_REGION = os.getenv("MINIO_REGION", "us-east-1")
MINIO_SECURE = os.getenv("MINIO_SECURE", "false").strip().lower() in (
    "1", "true", "yes", "on",
)
MINIO_PRESIGNED_EXPIRY = int(os.getenv("MINIO_PRESIGNED_EXPIRY", "3600"))

_s3_client = None
_s3_presign_client = None


def _is_minio_enabled() -> bool:
    return STORAGE_BACKEND == "minio"


def _normalize_endpoint(raw: str) -> str:
    url = raw.rstrip("/")
    if not url.startswith("http"):
        scheme = "https" if MINIO_SECURE else "http"
        url = f"{scheme}://{url}"
    return url


def _get_s3_client():
    global _s3_client
    if _s3_client is not None:
        return _s3_client

    import boto3
    from botocore.config import Config

    endpoint_url = _normalize_endpoint(MINIO_ENDPOINT)

    _s3_client = boto3.client(
        "s3",
        endpoint_url=endpoint_url,
        aws_access_key_id=MINIO_ACCESS_KEY,
        aws_secret_access_key=MINIO_SECRET_KEY,
        region_name=MINIO_REGION,
        config=Config(signature_version="s3v4"),
    )

    _ensure_bucket_exists()
    log.info(
        "MinIO client initialized — endpoint=%s bucket=%s",
        endpoint_url,
        MINIO_BUCKET,
    )
    return _s3_client


def _get_presign_client():
    """Return a client whose endpoint matches the public URL.

    generate_presigned_url is a local computation (no network call),
    so this client never needs to actually reach the public endpoint.
    The host baked into the signature will match what the browser sends.
    """
    global _s3_presign_client
    if _s3_presign_client is not None:
        return _s3_presign_client

    if not MINIO_PUBLIC_ENDPOINT:
        _s3_presign_client = _get_s3_client()
        return _s3_presign_client

    import boto3
    from botocore.config import Config

    public_url = _normalize_endpoint(MINIO_PUBLIC_ENDPOINT)

    _s3_presign_client = boto3.client(
        "s3",
        endpoint_url=public_url,
        aws_access_key_id=MINIO_ACCESS_KEY,
        aws_secret_access_key=MINIO_SECRET_KEY,
        region_name=MINIO_REGION,
        config=Config(signature_version="s3v4"),
    )
    log.info("Presign client initialized — public endpoint=%s", public_url)
    return _s3_presign_client


def _ensure_bucket_exists():
    from botocore.exceptions import ClientError

    client = _s3_client
    try:
        client.head_bucket(Bucket=MINIO_BUCKET)
    except ClientError:
        log.info("Bucket '%s' does not exist, creating it.", MINIO_BUCKET)
        client.create_bucket(Bucket=MINIO_BUCKET)


def upload_file(local_path: str, object_key: str) -> None:
    client = _get_s3_client()
    log.debug("Uploading %s -> s3://%s/%s", local_path, MINIO_BUCKET, object_key)
    client.upload_file(local_path, MINIO_BUCKET, object_key)


def upload_folder(folder_path: str) -> list[str]:
    """Upload every file in *folder_path* and return the list of object keys."""
    keys: list[str] = []
    folder_name = os.path.basename(folder_path)
    for root, _dirs, files in os.walk(folder_path):
        for fname in files:
            local = os.path.join(root, fname)
            rel = os.path.relpath(local, folder_path)
            key = f"{folder_name}/{rel}"
            upload_file(local, key)
            keys.append(key)
    return keys


def presigned_url(object_key: str, expiry: int | None = None) -> str:
    client = _get_presign_client()
    return client.generate_presigned_url(
        "get_object",
        Params={"Bucket": MINIO_BUCKET, "Key": object_key},
        ExpiresIn=expiry or MINIO_PRESIGNED_EXPIRY,
    )


def public_url_for(folder_path: str, filename: str) -> str:
    """Return a presigned download URL for a file inside *folder_path*."""
    folder_name = os.path.basename(folder_path)
    key = f"{folder_name}/{filename}"
    return presigned_url(key)


def delete_object(object_key: str) -> None:
    client = _get_s3_client()
    client.delete_object(Bucket=MINIO_BUCKET, Key=object_key)
    log.debug("Deleted s3://%s/%s", MINIO_BUCKET, object_key)


def delete_folder(folder_name: str) -> None:
    """Delete all objects under a folder prefix."""
    client = _get_s3_client()
    prefix = f"{folder_name}/"
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=MINIO_BUCKET, Prefix=prefix):
        for obj in page.get("Contents", []):
            client.delete_object(Bucket=MINIO_BUCKET, Key=obj["Key"])
            log.debug("Deleted s3://%s/%s", MINIO_BUCKET, obj["Key"])


def schedule_cleanup(folder_path: str, delay_minutes: int) -> None:
    """Schedule deletion of both local files and MinIO objects after a delay."""

    def _do_cleanup():
        time.sleep(delay_minutes * 60)

        folder_name = os.path.basename(folder_path)
        try:
            delete_folder(folder_name)
        except Exception as exc:
            log.error("Error cleaning up MinIO objects for %s: %s", folder_name, exc)

        try:
            import shutil
            shutil.rmtree(folder_path)
            log.debug("Local folder %s deleted.", folder_path)
        except Exception as exc:
            log.error("Error deleting local folder %s: %s", folder_path, exc)

    thread = threading.Thread(target=_do_cleanup, daemon=True)
    thread.start()
