# MsgKit MIME Compatibility

## Valid MIME and optional identifiers

RFC 2045 section 7 makes Content-ID generally optional. RFC 2183 section 2.1
uses `inline` to request automatic display; it does not require an HTML
`cid:` reference. RFC 2298 includes a read-receipt example whose notification
part has no Content-ID. Such parts are not inherently malformed.

An unmatched HTML `cid:` reference is a separate issue. RFC 2110 also allows
related parts to be identified by Content-Location. The original archived
message that triggered the constructor exception is still unavailable.

## Initial upstream comparison

Twelve synthetic EML cases were converted to MSG without Outlook. Checks
inspected attachment bytes, names, IDs, body content, and HTML references.
The inputs used period-compatible MIME structures and a 1998 date.

The local fork and upstream MsgKit 3.0.4 produced identical results, including
with upstream's MimeKit 4.16.0 dependency and the fork's MimeKit 4.18.2.
The comparison was subsequently repeated against upstream MsgKit 3.0.6.
Each configuration passed ten checks and failed two notification checks.

| Case | Unmodified converter result |
| --- | --- |
| Missing, blank, or unusable Content-ID | Attachment bytes retained |
| Missing filename | Bytes retained; fallback name supplied |
| Matching HTML cid | Reference, ID, and bytes retained |
| Unmatched HTML cid | Reference retained; no target ID invented |
| Relative Content-Location | Bytes/name/HTML retained; location tag absent |
| Read receipt with or without inline disposition | Notification data empty |

The notification tests verified that recipient and disposition fields were
present before conversion. The specialized MsgKit notification branch set a
filename but never wrote those fields into its attachment stream. This is a
data-preservation defect, separate from constructor validation.

The earlier constructor probe demonstrated that upstream rejects a direct
`Attachments.Add` call with `isInline: true` and a missing ID. The EML
converter avoids that condition for the examples above. Production
regressions therefore exercise conversion, rather than require the upstream
constructor to accept values it intentionally rejects.

## Email.ToolKit workaround

The library now uses public MsgKit 3.0.6 and MimeKit 4.18.2. `EmlConverter`
inspects the MIME body and replaces disposition-notification parts only in
an internal conversion copy. Their serialized fields become ordinary
attachments, defaulting to `read-receipt.txt`. Existing names and Content-IDs
are retained. No identifiers are generated.

The archived EML is never rewritten. Messages without notifications are
passed to MsgKit from the original input bytes. Attached messages remain
attached EML files; their internal MIME bodies are not normalized.

This preserves notification information as readable attachment data. It does
not create a native Outlook read-receipt object or restore unknown HTML
relationships. MIME inspection adds a parse to conversion; the inspected
message is disposed before MsgKit parses the selected input.

## Permanent regression coverage

`SourceCode/ToolKit.Tests/MsgKitConversionTests.cs` contains the original
conversion cases plus nested and standalone notifications, source-file
preservation, multiple receipts, existing metadata, and attached messages.
They run through the production adapter without Outlook.

The permanent baseline had five read-receipt failures. With the workaround,
all eighteen conversion cases pass against the public package. The valid
inline-constructor control also passes. No production build references the
fork project; its checkout is retained pending submodule cleanup.

Real Outlook import and rendering still require a release smoke test. These
synthetic cases do not establish compatibility with every archived message.
An upstream change is not required for the local workaround to function.

## Optional upstream report

`MsgKitReadReceiptIssue.md` contains an unsent issue draft and reproduction
steps. `Examples/ReadReceiptWithoutContentId.eml` is synthetic test input.
The report concerns empty output data, rather than the earlier argument over
whether inline attachments need Content-IDs. No maintainer was contacted.

Initial comparison source and TRX results remain in the ignored local
`artifacts/msgkit-examination/` directory and examination ZIP.

## References

- [MIME identifiers](https://www.rfc-editor.org/rfc/rfc2045.html#section-7)
- [Inline disposition](https://www.rfc-editor.org/rfc/rfc2183.html#section-2.1)
- [Read-receipt example](https://www.rfc-editor.org/rfc/rfc2298.html#section-9)
- [Content-Location](https://www.rfc-editor.org/rfc/rfc2110.html)
- [Upstream package](https://www.nuget.org/packages/MsgKit/3.0.6)
