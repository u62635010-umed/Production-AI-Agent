from __future__ import annotations

from io import BytesIO

from PIL import Image, ImageOps, UnidentifiedImageError


SUPPORTED_FORMATS = {"JPEG", "PNG"}
MAX_IMAGE_BYTES = 30 * 1024 * 1024
MAX_IMAGE_EDGE = 4096


def process_image(image_bytes: bytes, rotation_degrees: int = 0) -> tuple[bytes, str]:
    """Apply EXIF orientation and optional user rotation; return PNG bytes."""
    if not image_bytes:
        raise ValueError("The uploaded image is empty.")
    if len(image_bytes) > MAX_IMAGE_BYTES:
        raise ValueError("Image exceeds the 30 MB upload limit.")
    try:
        with Image.open(BytesIO(image_bytes)) as source:
            if source.format not in SUPPORTED_FORMATS:
                raise ValueError("Unsupported image. Upload a JPG, JPEG, or PNG file.")
            image = ImageOps.exif_transpose(source).convert("RGB")
            rotation = int(rotation_degrees) % 360
            if rotation not in {0, 90, 180, 270}:
                raise ValueError("Rotation must be 0, 90, 180, or 270 degrees.")
            if rotation:
                image = image.rotate(rotation, expand=True)
            if max(image.size) > MAX_IMAGE_EDGE:
                ratio = MAX_IMAGE_EDGE / max(image.size)
                size = (max(1, round(image.width * ratio)), max(1, round(image.height * ratio)))
                image = image.resize(size, Image.Resampling.LANCZOS)
            out = BytesIO()
            image.save(out, format="PNG", optimize=True)
            return out.getvalue(), "image/png"
    except UnidentifiedImageError as exc:
        raise ValueError("The uploaded file could not be read as an image.") from exc
