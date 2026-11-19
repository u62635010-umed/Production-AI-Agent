from io import BytesIO

import pytest
from PIL import Image
from pydantic import ValidationError

from src.config import Settings
from src.image_processor import process_image
from src.schemas import ProductionReport
from src.vision_agent import extract_report


def image_bytes():
    stream = BytesIO()
    Image.new("RGB", (80, 40), "white").save(stream, format="JPEG")
    return stream.getvalue()


def test_uploaded_image_is_processed_and_rotated():
    result, mime = process_image(image_bytes(), rotation_degrees=90)
    assert mime == "image/png"
    with Image.open(BytesIO(result)) as image:
        assert image.size == (40, 80)
        assert image.format == "PNG"


def test_unreadable_image_has_clear_error():
    with pytest.raises(ValueError, match="could not be read"):
        process_image(b"not an image")


def test_production_report_json_is_validated():
    report = ProductionReport.model_validate_json('{"production_date":"2026-05-01","machine_id":"M1","records":[]}')
    assert report.production_date == "2026-05-01"
    with pytest.raises(ValidationError):
        ProductionReport.model_validate_json('{"records":"not-a-list"}')


def test_gemini_schema_avoids_unsupported_additional_properties():
    import json

    from src.schemas import ProductionReport

    assert "additionalProperties" not in json.dumps(ProductionReport.model_json_schema())


def test_missing_api_key_is_reported_without_network_call():
    with pytest.raises(RuntimeError, match="GEMINI_API_KEY"):
        extract_report(b"image", settings=Settings(gemini_api_key=""))
