import argparse
import json
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker


ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH = (
    ROOT / "contracts" / "hybrid" / "v1" / "response.schema.json"
)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--result", required=True, type=Path)
    args = parser.parse_args()

    raw = args.result.read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        raise SystemExit("Engine result contains a forbidden UTF-8 BOM.")

    schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
    result = json.loads(raw.decode("utf-8", errors="strict"))
    Draft202012Validator.check_schema(schema)
    errors = sorted(
        Draft202012Validator(
            schema, format_checker=FormatChecker()
        ).iter_errors(result),
        key=lambda error: list(error.absolute_path),
    )
    if errors:
        details = "\n".join(
            f"{'/'.join(map(str, error.absolute_path))}: {error.message}"
            for error in errors
        )
        raise SystemExit(details)

    if result["status"] != "completed":
        raise SystemExit("Expected the emitted integration result to complete.")
    if not result["candidates"]:
        raise SystemExit("Expected at least one emitted integration candidate.")
    print(
        "PASS|Hybrid Engine Result|"
        f"{len(result['candidates'])} candidates match Contract 1.0."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
