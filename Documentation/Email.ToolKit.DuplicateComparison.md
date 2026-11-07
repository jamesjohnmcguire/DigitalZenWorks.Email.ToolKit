# Email.ToolKit Duplicate Comparison

This reference describes the mail-hash inputs in 1.14.0 and provides an
initial MAPI-to-Outlook mapping for clients and the native MAPI work.
It records current behavior, not a proposal to change comparison rules.

## What a hash represents

`OutlookItem.Hash` and `GetHashAsync` select properties through the relevant
mail, appointment, or contact wrapper. The selected buffers are concatenated
and hashed with SHA-256; the result is Base64 text. The usual path uses
`strict = false`. The instance property caches its first successful result;
create a fresh wrapper when examining changes to an item.

This compares selected Outlook data, not the original EML/MSG file bytes.
Matching visible subject/body text alone does not establish equal hashes.
Read state, reminders, categories, delivery times, custom properties, and
attachment metadata can also affect the result.

Hashes should not be treated as a permanent cross-version identifier or as
proof of C#/native equivalence. Property selection, encoding, normalization,
ordering, and serialization all matter. No algorithm/profile version is
currently embedded in the returned hash string.

## Current mail inputs

The source of truth is [OutlookMail.GetProperties][mail-source], together
with [OutlookItem's serializers][item-source]. This inventory follows the
binary hash path; the human-readable details output is not its byte format.

Buffers are appended in this order: actions, attachments, dates, enums,
strings, user properties, then a packed 16-bit boolean value. The common
combiner currently also allocates two additional trailing zero bytes.
These details matter when reproducing existing hashes in native code.

### Boolean bits, in order

```text
AlternateRecipientAllowed, AutoForwarded, AutoResolvedWinner,
DeleteAfterSubmit, IsMarkedAsTask, NoAging,
OriginatorDeliveryReportRequested, ReadReceiptRequested,
RecipientReassignmentProhibited, ReminderOverrideDefault,
ReminderPlaySound, ReminderSet, Saved, Sent, Submitted, UnRead
```

`UnRead` is included today. A read and an unread copy can therefore compare
differently. Changing that policy requires a separate reviewed behavior
change and regression tests.

### Dates, in order

```text
DeferredDeliveryTime, ExpiryTime, ReceivedTime, ReminderTime,
RetentionExpirationDate, SentOn, TaskCompletedDate, TaskDueDate,
TaskStartDate
```

Dates are serialized as concatenated round-trip (`O`) strings, then UTF-8.
They are not hashed as raw MAPI FILETIME structures. Some unavailable COM
properties have fallback values; see the implementation before porting.

### Enum values, in order

```text
BodyFormat, Class, Importance, MarkForDownload, Permission,
PermissionService, Sensitivity
```

The serializer writes integer values into a byte buffer. `Class` here is
Outlook's object-class enum; it is distinct from the `MessageClass` string.

### Strings, in order for the default hash

```text
BCC, BillingInformation, Body, Categories, CC, Companies,
ConversationTopic, FlagRequest, transport headers, HTMLBody,
MessageClass, Mileage, ReceivedByName, ReminderSoundFile,
ReplyRecipientNames, RetentionPolicyName, SenderEmailAddress,
SenderEmailType, SenderName, SentOnBehalfOfName, Subject, TaskSubject,
To, VotingOptions, VotingResponse
```

The strings are concatenated without field delimiters and encoded as UTF-8.
Mail recipients enter through the display strings `To`, `CC`, and `BCC`;
this path does not serialize each row of the `Recipients` collection.

The default path omits `ConversationID`. Strict string extraction also adds
`ReceivedByEntryID`, `ReceivedOnBehalfOfEntryID`, and
`ReceivedOnBehalfOfName`; ordinary hash entry points do not request it.
`strict` does not automatically include every omitted Outlook property.

### Collections and normalization

- Actions include `CopyLike`, `Enabled`, `Name`, `Prefix`, `ReplyStyle`,
  `ResponseStyle`, and `ShowOn`, in collection order.
- Attachments include `DisplayName`, `Index`, `Position`, `Type`, available
  filename/path metadata, and Base64-encoded saved content. Collection order
  matters. The implementation uses temporary files; it is not a raw hash of
  the MAPI attachment table.
- User properties include `Formula`, `Name`, `Type`, `ValidationFormula`,
  `ValidationText`, and formatted `Value`, in collection order.
- In non-strict mode, `Body` uses `TrimEnd`, `HTMLBody` passes through
  `HtmlEmail.Trim`, and headers remove the matched MimeOLE version suffix,
  normalize `Errors-to:` to `Errors-To:`, and sort header lines.
- Header splitting/joining has target-framework conditional branches, and
  sorting uses the default string comparer. Do not assume hashes are proven
  identical across runtimes, cultures, or Outlook representations.

### Selected exclusions

`EntryID`, `CreationTime`, `LastModificationTime`, `ConversationIndex`,
`Size`, `DownloadState`, and `RTFBody` are not directly selected by the mail
hash path. Neither is the item's parent folder or store identity. These
omissions allow some storage-specific differences to be ignored, but they
are not a documented rationale for every historical exclusion.

Properties may also appear indirectly in included transport headers or user
properties; excluding a direct Outlook property is not a guarantee that all
related information has been removed from the comparison data.

## Initial MAPI-to-Outlook mapping

The canonical names below follow Microsoft's [MAPI name mapping][mapi-map].
The Outlook column identifies the related object-model view; correspondence
alone does not establish byte-for-byte equivalence. This is a starting set,
not a complete inventory of mail, appointment, or contact properties.

| MAPI name | Canonical name | Outlook view / default mail hash |
| --- | --- | --- |
| PR_SUBJECT | PidTagSubject | Subject; included |
| PR_BODY | PidTagBody | Body; normalized |
| PR_MESSAGE_CLASS | PidTagMessageClass | MessageClass; included |
| PR_TRANSPORT_MESSAGE_HEADERS | PidTagTransportMessageHeaders | PropertyAccessor 0x007D001F; normalized |
| PR_DISPLAY_TO | PidTagDisplayTo | To display string; included |
| PR_DISPLAY_CC | PidTagDisplayCc | CC display string; included |
| PR_DISPLAY_BCC | PidTagDisplayBcc | BCC display string; included |
| PR_SENDER_NAME | PidTagSenderName | SenderName; included |
| PR_SENDER_EMAIL_ADDRESS | PidTagSenderEmailAddress | SenderEmailAddress; included |
| PR_SENDER_ADDRTYPE | PidTagSenderAddressType | SenderEmailType; included |
| PR_CLIENT_SUBMIT_TIME | PidTagClientSubmitTime | SentOn; included |
| PR_MESSAGE_DELIVERY_TIME | PidTagMessageDeliveryTime | ReceivedTime; included |
| PR_IMPORTANCE | PidTagImportance | Importance; included |
| PR_SENSITIVITY | PidTagSensitivity | Sensitivity; included |
| PR_ENTRYID | PidTagEntryId | EntryID; excluded directly |
| PR_CREATION_TIME | PidTagCreationTime | CreationTime; excluded directly |
| PR_LAST_MODIFICATION_TIME | PidTagLastModificationTime | LastModificationTime; excluded directly |
| PR_MESSAGE_ATTACHMENTS | PidTagMessageAttachments | Attachment table/collection; requires per-attachment mapping |
| PR_MESSAGE_RECIPIENTS | PidTagMessageRecipients | Recipient table/collection; not a serialized mail-hash input |

Microsoft documents [MessageClass][oom-messageclass] as corresponding to
PidTagMessageClass. Do not map it to [Class][oom-class], which returns an
OlObjectClass enum. This corrects an ambiguity in the old Properties notes.

Similarly, display-recipient strings do not replace a complete recipient
row mapping. The [To property][oom-to] describes display names; native
address resolution and recipient-table traversal need separate decisions.

For named properties, record the property-set GUID and numeric/string name,
plus data type. A store-assigned property ID alone is not a portable key;
see Microsoft's [named-property guidance][named-properties]. Expand the
mapping with those identities rather than guessing fixed tags for task,
reminder, category, contact, or appointment fields.

For each added mapping, record its item type, canonical and MAPI names,
property type/tag or named-property identity, object-model accessor,
inclusion rule, normalization, evidence, and any provider-specific caveats.
Unknown or unverified mappings should remain explicitly labelled.

## Appointment/contact and native follow-up

The dispatcher also supports appointment and contact wrappers, which have
their own selected properties. Appointments include a separate recipient
buffer; contacts and mail do not. Their complete field-level inventories
and mappings remain follow-up work, not implied by the mail table above.

The local `mapi-library` branch contains a prototype in
`SourceCode/MapiLibrary/Message.cpp`, inspected at commit `8466d1b`.
It requests a different property list and its hash loop currently handles
string-valued properties while logging/skipping other returned types.
Its encoding/buffer construction and output representation need review
before comparing its hashes with C#. No native/C# parity is established.

The original `MapiProperties.txt`, `Properties.txt`, and
`MapiCompareMessagesSnippets.txt` remain intact in the adjacent Email
repository's Support directory as research inputs. Their included/ignored
labels and comparison snippets are not automatically the current policy.

## Changing comparison rules

Before changing selection or normalization:

1. Record the existing and intended included/excluded properties, reason,
   affected item types, and version here.
2. Add cases that should compare equal and cases that must stay distinct,
   including attachments, read state, dates, recipient representations, and
   missing properties as relevant to the change.
3. Check property order, encoding, culture/time conversion, missing values,
   and delimiters as well as property identity. A shared property list alone
   cannot guarantee equal serialized data.
4. Update client release notes. If native and C# hashes must agree, establish
   shared fixtures and explicit serialization/profile rules first.

| Release | Documented policy change |
| --- | --- |
| 1.14.0 baseline | Initial source-derived mail inventory recorded; no hashing behavior changed by this documentation work. |

Remaining work is tracked in [Email.ToolKit.Tasks.md](Email.ToolKit.Tasks.md).

[mail-source]: ../SourceCode/ToolKit.Library/OutlookMail.cs
[item-source]: ../SourceCode/ToolKit.Library/OutlookItem.cs
[mapi-map]: https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapping-mapi-names-to-canonical-property-names
[oom-messageclass]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._mailitem.messageclass?view=outlook-pia
[oom-class]: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.class
[oom-to]: https://learn.microsoft.com/en-us/office/vba/api/outlook.mailitem.to
[named-properties]: https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/about-named-properties-used-by-outlook
