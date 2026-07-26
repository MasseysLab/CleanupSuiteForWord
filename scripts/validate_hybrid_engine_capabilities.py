import argparse
import json
import subprocess
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker


ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH = (
    ROOT / "contracts" / "hybrid" / "v1" / "capabilities.schema.json"
)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--engine", required=True, type=Path)
    args = parser.parse_args()

    completed = subprocess.run(
        [str(args.engine), "--capabilities"],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        timeout=5,
    )
    if completed.returncode != 0:
        raise SystemExit(
            f"Capability command failed with exit code {completed.returncode}."
        )
    if completed.stderr:
        raise SystemExit("Capability command wrote unexpected diagnostic output.")

    schema = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
    capability = json.loads(completed.stdout)
    Draft202012Validator.check_schema(schema)
    errors = sorted(
        Draft202012Validator(
            schema, format_checker=FormatChecker()
        ).iter_errors(capability),
        key=lambda error: list(error.absolute_path),
    )
    if errors:
        details = "\n".join(
            f"{'/'.join(map(str, error.absolute_path))}: {error.message}"
            for error in errors
        )
        raise SystemExit(details)

    security = capability["security"]
    unsafe = [
        name
        for name in (
            "editsWordDocuments",
            "requiresNetwork",
            "opensListeningEndpoint",
            "runsAsService",
            "requiresElevation",
            "logsDocumentContent",
        )
        if security[name]
    ]
    if unsafe:
        raise SystemExit(
            "Capability handshake contains unsafe flags: " + ", ".join(unsafe)
        )

    print(
        "PASS|Hybrid Engine Capabilities|"
        f"{capability['engine']['id']} {capability['engine']['version']} "
        "matches Contract 1.0."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
