import warnings
from io import BytesIO
from pathlib import Path

from django.core.files.base import ContentFile
from PIL import Image, ImageOps, UnidentifiedImageError
from pillow_heif import register_heif_opener



MAX_PHOTOS_PER_UPLOAD = 25
MAX_PHOTO_SIZE = 15 * 1024 * 1024
MAX_TOTAL_UPLOAD_SIZE = 150 * 1024 * 1024

ALLOWED_IMAGE_FORMATS = {
    "JPEG",
    "PNG",
    "WEBP",
    "HEIC",
    "HEIF",
    "MPO",
}

HEIF_IMAGE_FORMATS = {
    "HEIC",
    "HEIF",
    "MPO",
}

Image.MAX_IMAGE_PIXELS = 60_000_000


class PhotoUploadError(ValueError):
    pass


def validate_photo_uploads(uploads):
    uploads = list(uploads)

    if not uploads:
        raise PhotoUploadError(
            "Choose at least one photo."
        )

    if len(uploads) > MAX_PHOTOS_PER_UPLOAD:
        raise PhotoUploadError(
            "Upload no more than "
            f"{MAX_PHOTOS_PER_UPLOAD} photos at once."
        )

    total_size = sum(
        upload.size or 0
        for upload in uploads
    )

    if total_size > MAX_TOTAL_UPLOAD_SIZE:
        raise PhotoUploadError(
            "The selected photos are too large together. "
            "Choose fewer photos and try again."
        )

    for upload in uploads:
        if (upload.size or 0) > MAX_PHOTO_SIZE:
            raise PhotoUploadError(
                f"{upload.name} is larger than 15 MB."
            )

        try:
            upload.seek(0)

            with warnings.catch_warnings():
                warnings.simplefilter(
                    "error",
                    Image.DecompressionBombWarning,
                )

                with Image.open(upload) as image:
                    image_format = (
                        image.format or ""
                    ).upper()

                    image.verify()

            if image_format not in ALLOWED_IMAGE_FORMATS:
                raise PhotoUploadError(
                    f"{upload.name} is not a supported photo "
                    f"(detected format: {image_format or 'unknown'})."
                )

        except PhotoUploadError:
            raise

        except (
            UnidentifiedImageError,
            OSError,
            ValueError,
            Image.DecompressionBombError,
            Image.DecompressionBombWarning,
        ) as exc:
            raise PhotoUploadError(
                f"{upload.name} is not a valid photo."
            ) from exc

        finally:
            upload.seek(0)

    return uploads


def prepare_photo_for_storage(upload):
    """Convert HEIC/HEIF uploads to browser-compatible JPEG files."""
    try:
        upload.seek(0)

        with Image.open(upload) as image:
            image_format = (
                image.format or ""
            ).upper()

            if image_format not in HEIF_IMAGE_FORMATS:
                return upload

            image = ImageOps.exif_transpose(image)

            if image.mode in {"RGBA", "LA"}:
                image_with_alpha = image.convert("RGBA")
                jpeg_image = Image.new(
                    "RGB",
                    image_with_alpha.size,
                    "white",
                )
                jpeg_image.paste(
                    image_with_alpha,
                    mask=image_with_alpha.getchannel("A"),
                )
            else:
                jpeg_image = image.convert("RGB")

            output = BytesIO()
            jpeg_image.save(
                output,
                format="JPEG",
                quality=92,
                optimize=True,
            )

        output.seek(0)

        return ContentFile(
            output.getvalue(),
            name=f"{Path(upload.name).stem}.jpg",
        )

    except (
        UnidentifiedImageError,
        OSError,
        ValueError,
        Image.DecompressionBombError,
        Image.DecompressionBombWarning,
    ) as exc:
        raise PhotoUploadError(
            f"{upload.name} could not be converted "
            "to a browser-compatible photo."
        ) from exc

    finally:
        upload.seek(0)


def create_photo_records(
    photo_model,
    parent_field,
    parent,
    uploads,
):
    saved_files = []

    try:
        for upload in uploads:
            stored_upload = prepare_photo_for_storage(upload)

            photo = photo_model(
                **{
                    parent_field: parent,
                    "image": stored_upload,
                }
            )

            try:
                photo.save()
            except Exception:
                if (
                    photo.image
                    and photo.image.name
                    and getattr(
                        photo.image,
                        "_committed",
                        False,
                    )
                ):
                    saved_files.append((
                        photo.image.storage,
                        photo.image.name,
                    ))
                raise

            saved_files.append((
                photo.image.storage,
                photo.image.name,
            ))

    except Exception:
        for storage, name in saved_files:
            try:
                storage.delete(name)
            except Exception:
                pass

        raise

    return len(saved_files)
