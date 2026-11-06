# Draft: ConvertEmlToMsg discards disposition-notification fields

This is a local issue draft. It has not been submitted upstream.

## Problem

Converting a multipart/report containing a message/disposition-notification
part creates an MSG attachment with an empty data stream. The input includes
Final-Recipient, Original-Message-ID, and Disposition fields. The human-readable
message body converts successfully, but the machine-readable fields are lost.

This reproduces with public MsgKit 3.0.6. It also reproduced with 3.0.4 and
with the local attachment-tolerance fork. No Outlook installation or COM
activation is needed to reproduce the conversion defect.

## Input and reproduction

Use [the synthetic EML](Examples/ReadReceiptWithoutContentId.eml). Its missing
Content-ID is intentional: RFC 2298 section 9 includes the same general MIME
structure without a Content-ID on the notification part.

Create a console project targeting .NET 10 and reference MsgKit 3.0.6,
MimeKit 4.18.2, OpenMcdf 3.3.0, and System.Security.Cryptography.Pkcs 10.0.13.
Copy the EML into the working directory and run:

```csharp
using System;
using MsgKit;
using OpenMcdf;

Converter.ConvertEmlToMsg(
    "ReadReceiptWithoutContentId.eml", "receipt.msg");

using RootStorage root = RootStorage.OpenRead("receipt.msg");
Storage attachment = root.OpenStorage("__attach_version1.0_#00000000");
using CfbStream data = attachment.OpenStream(
    PropertyTags.PR_ATTACH_DATA_BIN.Name);
Console.WriteLine(data.Length);
```

Start with no existing `receipt.msg`; the converter creates a new file.
Expected: a nonempty attachment containing the notification fields.
Actual: the printed attachment length is zero.

## Apparent cause

In `Converter.ConvertEmlToMsg`, this branch sets the filename and exits:

```csharp
case MessageDispositionNotification notification:
    fileName = notification.FileName;
    break;
```

Unlike the delivery-status branch, it does not write the notification body
to `attachmentStream`. One possible upstream correction is to serialize its
content before exiting, then verify the decoded MSG attachment fields.

The regression should assert retained recipient, original-message ID, and
disposition values, both with and without an inline disposition header.
Checking only that conversion does not throw would miss the data loss.

## Scope

This report concerns preservation of parsed notification data. It does not
request a change to attachment-constructor validation, or assume every
notification is referenced by an HTML cid URL.

Email.ToolKit has a local workaround that represents these fields as an
ordinary attachment in a conversion copy. An upstream correction would let
that workaround be retired once a fixed package is available and tested.

- [RFC 2298 example](https://www.rfc-editor.org/rfc/rfc2298.html#section-9)
- [MsgKit package](https://www.nuget.org/packages/MsgKit/3.0.6)
- [Current converter source](https://github.com/Sicos1977/MsgKit/blob/master/MsgKit/Converter.cs)
