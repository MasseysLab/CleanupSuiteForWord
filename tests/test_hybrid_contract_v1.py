import copy
import hashlib
import json
import unittest
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker


ROOT = Path(__file__).resolve().parents[1]
CONTRACT = ROOT / "contracts" / "hybrid" / "v1"
FIXTURES = CONTRACT / "fixtures"


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def sha256_utf8(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def utf16_length(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


def utf16_slice(text: str, start: int, end: int) -> str:
    encoded = text.encode("utf-16-le")
    return encoded[start * 2 : end * 2].decode("utf-16-le")


class HybridContractV1Tests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.schemas = {
            "request": load_json(CONTRACT / "request.schema.json"),
            "response": load_json(CONTRACT / "response.schema.json"),
            "capabilities": load_json(CONTRACT / "capabilities.schema.json"),
            "manifest": load_json(
                CONTRACT / "installation-manifest.schema.json"
            ),
        }
        cls.validators = {}
        for name, schema in cls.schemas.items():
            Draft202012Validator.check_schema(schema)
            cls.validators[name] = Draft202012Validator(
                schema, format_checker=FormatChecker()
            )

        cls.request = load_json(FIXTURES / "request.valid.json")
        cls.response = load_json(FIXTURES / "response.valid.json")
        cls.error_response = load_json(FIXTURES / "response.error.valid.json")
        cls.structural_response = load_json(
            FIXTURES / "response.structural.valid.json"
        )
        cls.capabilities = load_json(FIXTURES / "capabilities.valid.json")
        cls.manifest = load_json(
            FIXTURES / "installation-manifest.valid.json"
        )
        cls.snapshot = load_json(FIXTURES / "snapshot.json")
        cls.supplementary_snapshot = load_json(
            FIXTURES / "snapshot.supplementary-unicode.json"
        )
        cls.protocol = load_json(CONTRACT / "protocol.json")
        cls.vocabulary = load_json(
            CONTRACT / "operation-vocabulary.json"
        )

    def assertValid(self, validator_name, instance):
        errors = sorted(
            self.validators[validator_name].iter_errors(instance),
            key=lambda error: list(error.absolute_path),
        )
        if errors:
            messages = [
                f"{'/'.join(map(str, error.absolute_path))}: {error.message}"
                for error in errors
            ]
            self.fail("\n".join(messages))

    def assertInvalid(self, validator_name, instance):
        self.assertTrue(
            list(self.validators[validator_name].iter_errors(instance)),
            f"Expected {validator_name} instance to be rejected.",
        )

    def test_machine_readable_schemas_and_valid_fixtures(self):
        self.assertValid("request", self.request)
        self.assertValid("response", self.response)
        self.assertValid("response", self.error_response)
        self.assertValid("response", self.structural_response)
        self.assertValid("capabilities", self.capabilities)
        self.assertValid("manifest", self.manifest)

    def test_snapshot_is_exact_utf8_without_newline_normalization(self):
        text = self.snapshot["text"]
        self.assertIn("\r", text)
        self.assertNotIn("\n", text)
        self.assertEqual(
            len(text.encode("utf-8")), self.snapshot["byteLength"]
        )
        self.assertEqual(utf16_length(text), self.snapshot["utf16Length"])
        self.assertEqual(sha256_utf8(text), self.snapshot["sha256"])
        self.assertEqual(
            self.request["snapshot"]["sha256"], self.snapshot["sha256"]
        )
        self.assertEqual(
            self.request["snapshot"]["snapshotId"], self.snapshot["sha256"]
        )

    def test_supplementary_unicode_uses_word_utf16_range_positions(self):
        fixture = self.supplementary_snapshot
        text = fixture["text"]
        self.assertEqual(len(text.encode("utf-8")), fixture["byteLength"])
        self.assertEqual(utf16_length(text), fixture["utf16Length"])
        self.assertEqual(sha256_utf8(text), fixture["sha256"])
        emoji = utf16_slice(
            text,
            fixture["emojiStartUtf16"],
            fixture["emojiEndUtf16"],
        )
        self.assertEqual(emoji, "😀")
        self.assertEqual(utf16_length(emoji), 2)
        self.assertEqual(sha256_utf8(emoji), fixture["emojiSha256"])

    def test_request_scope_matches_snapshot_and_excludes_identity(self):
        scope = self.request["scope"]
        snapshot = self.request["snapshot"]
        self.assertLessEqual(scope["startUtf16"], scope["endUtf16"])
        self.assertEqual(
            scope["endUtf16"] - scope["startUtf16"],
            snapshot["utf16Length"],
        )
        self.assertFalse(self.request["privacy"]["documentPathIncluded"])
        self.assertFalse(self.request["privacy"]["documentNameIncluded"])

        forbidden_keys = {
            "documentPath",
            "documentName",
            "url",
            "absolutePath",
        }

        def visit(value):
            if isinstance(value, dict):
                self.assertTrue(forbidden_keys.isdisjoint(value.keys()))
                for child in value.values():
                    visit(child)
            elif isinstance(value, list):
                for child in value:
                    visit(child)

        visit(self.request)

    def test_response_echo_and_candidate_fingerprints_match_snapshot(self):
        self.assertEqual(self.response["jobId"], self.request["jobId"])
        self.assertEqual(
            self.response["echo"]["snapshotId"],
            self.request["snapshot"]["snapshotId"],
        )
        self.assertEqual(
            self.response["echo"]["toolId"], self.request["tool"]["id"]
        )
        self.assertEqual(
            self.response["echo"]["toolDefinitionVersion"],
            self.request["tool"]["definitionVersion"],
        )

        text = self.snapshot["text"]
        candidate = self.response["candidates"][0]
        location = candidate["location"]
        fingerprint = candidate["fingerprint"]
        start = location["startUtf16"]
        end = location["endUtf16"]
        exact = utf16_slice(text, start, end)
        prefix_length = fingerprint["prefixLengthUtf16"]
        suffix_length = fingerprint["suffixLengthUtf16"]
        prefix = utf16_slice(text, max(0, start - prefix_length), start)
        suffix = utf16_slice(text, end, end + suffix_length)

        self.assertEqual(
            sha256_utf8(exact), fingerprint["exactTextSha256"]
        )
        self.assertEqual(
            sha256_utf8(prefix), fingerprint["prefixSha256"]
        )
        self.assertEqual(
            sha256_utf8(suffix), fingerprint["suffixSha256"]
        )
        self.assertEqual(
            fingerprint["snapshotId"], self.snapshot["sha256"]
        )

    def test_candidates_summaries_and_operations_are_coherent(self):
        candidates = self.response["candidates"]
        candidate_ids = [
            candidate["candidateId"] for candidate in candidates
        ]
        self.assertEqual(len(candidate_ids), len(set(candidate_ids)))
        snapshot_length = self.request["snapshot"]["utf16Length"]

        counts = {
            "applicable": 0,
            "review-only": 0,
            "protected": 0,
            "skipped": 0,
        }
        vocabulary = {
            operation["id"]: operation
            for operation in self.vocabulary["operations"]
        }

        for candidate in candidates:
            location = candidate["location"]
            self.assertLessEqual(
                location["startUtf16"], location["endUtf16"]
            )
            self.assertLessEqual(
                location["endUtf16"], snapshot_length
            )
            counts[candidate["state"]] += 1

            operation = candidate["operation"]
            definition = vocabulary[operation["type"]]
            self.assertEqual(
                operation["safetyClass"], definition["safetyClass"]
            )
            for parameter in definition["requiredParameters"]:
                self.assertIn(parameter, operation["parameters"])

            if candidate["state"] == "applicable":
                self.assertNotEqual(operation["type"], "reportOnly")
                self.assertTrue(
                    candidate["revalidation"][
                        "requireWholeScopeSnapshot"
                    ]
                )
                self.assertTrue(
                    candidate["revalidation"]["requireExactRange"]
                )
                self.assertTrue(
                    candidate["revalidation"]["requireContext"]
                )
                self.assertEqual(
                    candidate["revalidation"]["onMismatch"],
                    "abort-apply",
                )
                self.assertFalse(
                    candidate["revalidation"]["allowRelocation"]
                )
            else:
                self.assertEqual(operation["type"], "reportOnly")

            if operation["safetyClass"] == "structural-destructive":
                self.assertTrue(
                    candidate["revalidation"]["requireStructure"]
                )
                self.assertIn(
                    "structureSha256", candidate["fingerprint"]
                )

        summary = self.response["summary"]
        self.assertEqual(summary["total"], len(candidates))
        self.assertEqual(summary["applicable"], counts["applicable"])
        self.assertEqual(summary["reviewOnly"], counts["review-only"])
        self.assertEqual(summary["protected"], counts["protected"])
        self.assertEqual(summary["skipped"], counts["skipped"])

    def test_operation_schema_enum_matches_vocabulary(self):
        schema_operations = set(
            self.schemas["response"]["$defs"]["candidate"]["properties"][
                "operation"
            ]["properties"]["type"]["enum"]
        )
        vocabulary_operations = {
            operation["id"]
            for operation in self.vocabulary["operations"]
        }
        self.assertEqual(schema_operations, vocabulary_operations)

    def test_protocol_is_fail_closed_and_exit_codes_are_unambiguous(self):
        safety = self.protocol["applySafety"]
        self.assertFalse(safety["engineMayEditWord"])
        self.assertTrue(safety["validateWholeScopeBeforePreview"])
        self.assertTrue(safety["validateWholeScopeBeforeApply"])
        self.assertTrue(
            safety["validateAllCandidatesBeforeFirstMutation"]
        )
        self.assertEqual(
            safety["defaultMismatchAction"], "abort-apply"
        )
        self.assertFalse(safety["candidateRelocationAllowed"])
        exit_codes = list(self.protocol["exitCodes"].values())
        self.assertEqual(len(exit_codes), len(set(exit_codes)))
        failure = self.protocol["failureBehavior"]
        self.assertFalse(failure["partialResultsAllowed"])
        self.assertFalse(failure["reuseResultsAfterFailure"])
        self.assertFalse(
            failure["silentBehaviorChangingFallbackAllowed"]
        )
        self.assertEqual(
            failure["staleCandidateAction"],
            "abort-apply-before-first-mutation",
        )

    def test_capabilities_are_offline_nonprivileged_and_content_safe(self):
        security = self.capabilities["security"]
        self.assertFalse(security["editsWordDocuments"])
        self.assertFalse(security["requiresNetwork"])
        self.assertFalse(security["opensListeningEndpoint"])
        self.assertFalse(security["runsAsService"])
        self.assertFalse(security["requiresElevation"])
        self.assertFalse(security["logsDocumentContent"])
        self.assertTrue(
            self.capabilities["distribution"][
                "authenticodeRequiredForOfficialBeta"
            ]
        )
        protocol_minimum = tuple(
            map(int, self.capabilities["protocolRange"]["minimum"].split("."))
        )
        protocol_maximum = tuple(
            map(int, self.capabilities["protocolRange"]["maximum"].split("."))
        )
        requested_protocol = tuple(
            map(int, self.request["client"]["protocolVersion"].split("."))
        )
        self.assertLessEqual(protocol_minimum, requested_protocol)
        self.assertLessEqual(requested_protocol, protocol_maximum)
        supported_tool = self.capabilities["supportedTools"][0]
        self.assertEqual(supported_tool["id"], self.request["tool"]["id"])
        self.assertIn(
            self.request["tool"]["definitionVersion"],
            supported_tool["definitionVersions"],
        )
        self.assertIn(
            self.request["tool"]["analysisMode"],
            supported_tool["analysisModes"],
        )

    def test_installation_manifest_binds_template_and_engine(self):
        self.assertTrue(self.manifest["authenticodeRequired"])
        components = {
            component["id"]: component
            for component in self.manifest["components"]
        }
        self.assertIn("word-template", components)
        self.assertIn("analysis-engine", components)
        for component in components.values():
            relative_path = component["relativePath"]
            self.assertNotIn("..", relative_path)
            self.assertFalse(relative_path.startswith(("\\", "/")))
            self.assertNotRegex(relative_path, r"^[A-Za-z]:")
            self.assertRegex(
                component["sha256"], r"^[0-9a-f]{64}$"
            )
        self.assertEqual(
            components["analysis-engine"]["version"],
            self.capabilities["engine"]["version"],
        )

    def test_schema_rejects_relaxed_safety_and_untrusted_paths(self):
        relocated = copy.deepcopy(self.response)
        relocated["candidates"][0]["revalidation"][
            "allowRelocation"
        ] = True
        self.assertInvalid("response", relocated)

        content_log = copy.deepcopy(self.response)
        content_log["diagnostics"][
            "logContainsDocumentContent"
        ] = True
        self.assertInvalid("response", content_log)

        failed_with_candidates = copy.deepcopy(self.error_response)
        failed_with_candidates["candidates"] = copy.deepcopy(
            self.response["candidates"]
        )
        self.assertInvalid("response", failed_with_candidates)

        absolute_manifest = copy.deepcopy(self.manifest)
        absolute_manifest["components"][1][
            "relativePath"
        ] = "C:\\Program Files\\CleanupSuite.Engine.exe"
        self.assertInvalid("manifest", absolute_manifest)

        missing_engine = copy.deepcopy(self.manifest)
        missing_engine["components"][1]["id"] = "rules"
        self.assertInvalid("manifest", missing_engine)

        extra_request_field = copy.deepcopy(self.request)
        extra_request_field["documentPath"] = "C:\\private.docx"
        self.assertInvalid("request", extra_request_field)

    def test_error_response_never_returns_candidates(self):
        self.assertNotEqual(self.error_response["status"], "completed")
        self.assertEqual(self.error_response["candidates"], [])
        self.assertIn("error", self.error_response)

    def test_structural_candidate_requires_structure_fingerprint(self):
        candidate = self.structural_response["candidates"][0]
        self.assertEqual(
            candidate["operation"]["safetyClass"],
            "structural-destructive",
        )
        self.assertTrue(candidate["revalidation"]["requireStructure"])
        self.assertIn("structureSha256", candidate["fingerprint"])

        missing_structure = copy.deepcopy(self.structural_response)
        del missing_structure["candidates"][0]["fingerprint"][
            "structureSha256"
        ]
        self.assertInvalid("response", missing_structure)

        relaxed_structure = copy.deepcopy(self.structural_response)
        relaxed_structure["candidates"][0]["revalidation"][
            "requireStructure"
        ] = False
        self.assertInvalid("response", relaxed_structure)


if __name__ == "__main__":
    unittest.main()
