from __future__ import annotations

from google import genai
from google.genai import types

from src.config import Settings
from src.schemas import ProductionReport


EXTRACTION_PROMPT = """Read this photographed HPDC daily production report and return only structured data matching the supplied schema.

This photographed form is a three-shift HPDC Daily Production Report. Its general fields are Machine ID, Date, Component Name, No. of Cavity, Target, and Operator Name. It has three shift blocks: Shift 1 (7:00 AM to 3:30 PM), Shift 2 (3:30 PM to 11:00 PM), and Shift 3 (11:00 PM to 7:00 AM). Each shift has one Heat No. header, eight numbered hourly rows (1st through 8th), a TOTAL row, Remarks, Opt. Sign (Checked By), and Shift I/C Sign (Verified By).

Return shift-header values in shifts, one entry for each visible shift. For each filled hourly row return one records entry with record_type=hourly, shift_number, and hour_number from 1 through 8. The hourly columns are Shot Count, Total Shot, OK Shot, Rej. Shot, Type of Rej., Type of BD, and BD/no. of min. Keep shot counter, total shots, OK shots, rejected shots, produced parts if explicitly shown, rejection type/code, breakdown type/code, and breakdown minutes separate. Represent the TOTAL row as one record with record_type=shift_total. Include a shift_total row only when the printed TOTAL row has entered data; do not add row totals to the hourly lines.

Use only handwritten/filled production entries. The printed Type of Rejection code/reason table at the right and the printed Breakdown code/reason table at the bottom are reference legends, not measurements. If a handwritten rejection or breakdown code is visible in an hourly cell, preserve that exact code in the matching code field; a reference legend may help describe it but must never be returned as a production entry. Do not infer, calculate, or invent values. Keep shots, good shots, rejected shots, and produced parts as separate quantities. Never multiply shots by cavities.

For a clearly written zero, return numeric 0. For a genuinely blank numeric field, return null and list its exact field path in missing_fields. For unreadable handwriting, return null and list its exact field path and what is unclear in review_required. Do not treat an unclear entry as blank or zero. If the form's field labels or row relationships are unclear, describe the uncertainty in review_required. Return null for fields not present. Use ISO YYYY-MM-DD when a full date is readable; otherwise preserve the written date and flag it if the year or value is unclear.

Include field paths such as production_date, number_of_cavities, target, shifts[0].heat_number, records[0].hour_number, records[0].total_shots, and records[0].rejected_shots. Ignore preprinted reference-code tables as production data. Do not force values into a field whose unit or meaning is unclear; use review_required and other_entries where appropriate."""


def extract_report(image_bytes: bytes, mime_type: str = "image/png", settings: Settings | None = None) -> ProductionReport:
    settings = settings or Settings()
    if not settings.gemini_api_key:
        raise RuntimeError("Gemini API key is missing. Set GEMINI_API_KEY in .env to extract a real report photo.")
    if not settings.gemini_model:
        raise RuntimeError("GEMINI_MODEL is empty. Set it to a supported Gemini Flash model in .env.")
    client = genai.Client(
        api_key=settings.gemini_api_key,
        http_options=types.HttpOptions(
            timeout=60_000,
            retryOptions=types.HttpRetryOptions(attempts=1),
        ),
    )
    try:
        response = client.models.generate_content(
            model=settings.gemini_model,
            contents=[
                EXTRACTION_PROMPT,
                types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            ],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=ProductionReport,
                temperature=0.1,
            ),
        )
        if not response.text:
            raise RuntimeError("Gemini returned no structured response. Check the image and API/model access.")
        return ProductionReport.model_validate_json(response.text)
    except Exception as exc:
        message = str(exc)
        if "404" in message or "not found" in message.lower() or "unsupported" in message.lower():
            raise RuntimeError(f"Gemini model '{settings.gemini_model}' is unavailable for this API key. Change GEMINI_MODEL in .env.") from exc
        if isinstance(exc, RuntimeError):
            raise
        raise RuntimeError(f"Gemini extraction failed: {message}") from exc
