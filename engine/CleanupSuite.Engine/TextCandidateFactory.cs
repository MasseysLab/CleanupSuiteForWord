namespace MasseysLab.CleanupSuite.Engine;

public static class TextCandidateFactory
{
    private const int ContextLimitUtf16 = 32;

    public static CandidateRecord CreateReplacement(
        int ordinal,
        ValidatedJob job,
        int start,
        int end,
        string replacement,
        string reasonCode,
        string displayLabel)
    {
        string text = job.SnapshotText;
        int prefixStart = AlignedPrefixStart(text, start);
        int suffixEnd = AlignedSuffixEnd(text, end);
        string exact = text[start..end];
        string prefix = text[prefixStart..start];
        string suffix = text[end..suffixEnd];
        string paragraph = ParagraphContaining(text, start, end);

        return new CandidateRecord
        {
            CandidateId = $"c{ordinal:000000}",
            State = "applicable",
            ReasonCode = reasonCode,
            Location = new LocationRecord
            {
                StoryType = job.Request.Scope.StoryType,
                StartUtf16 = start,
                EndUtf16 = end,
                OffsetBasis = "snapshot-relative-utf16-code-units"
            },
            Fingerprint = new FingerprintRecord
            {
                Algorithm = "sha256-utf8-exact",
                SnapshotId = job.Request.Snapshot.SnapshotId,
                ExactTextSha256 = Hashing.Sha256Utf8(exact),
                PrefixLengthUtf16 = prefix.Length,
                PrefixSha256 = Hashing.Sha256Utf8(prefix),
                SuffixLengthUtf16 = suffix.Length,
                SuffixSha256 = Hashing.Sha256Utf8(suffix),
                ParagraphSha256 = Hashing.Sha256Utf8(paragraph)
            },
            Operation = new OperationRecord
            {
                Type = "replaceText",
                SafetyClass = "textual",
                Parameters = new Dictionary<string, object?>
                {
                    ["replacementText"] = replacement
                }
            },
            Revalidation = new RevalidationRecord
            {
                RequireWholeScopeSnapshot = true,
                RequireExactRange = true,
                RequireContext = true,
                RequireStructure = false,
                OnMismatch = "abort-apply",
                AllowRelocation = false
            },
            Confidence = new ConfidenceRecord
            {
                Tier = "deterministic",
                Score = 1.0m
            },
            Display = new DisplayRecord
            {
                Label = displayLabel,
                CanHighlight = true
            }
        };
    }

    private static int AlignedPrefixStart(string text, int candidateStart)
    {
        int start = Math.Max(0, candidateStart - ContextLimitUtf16);
        if (start > 0
            && start < candidateStart
            && char.IsLowSurrogate(text[start])
            && char.IsHighSurrogate(text[start - 1]))
        {
            start++;
        }

        return start;
    }

    private static int AlignedSuffixEnd(string text, int candidateEnd)
    {
        int end = Math.Min(
            text.Length,
            candidateEnd + ContextLimitUtf16);
        if (end > candidateEnd
            && end < text.Length
            && char.IsLowSurrogate(text[end])
            && char.IsHighSurrogate(text[end - 1]))
        {
            end--;
        }

        return end;
    }

    private static string ParagraphContaining(
        string text,
        int candidateStart,
        int candidateEnd)
    {
        int paragraphStart = candidateStart == 0
            ? 0
            : text.LastIndexOf('\r', candidateStart - 1) + 1;
        int paragraphMark = text.IndexOf('\r', candidateEnd);
        int paragraphEnd = paragraphMark < 0
            ? text.Length
            : paragraphMark + 1;
        return text[paragraphStart..paragraphEnd];
    }
}
