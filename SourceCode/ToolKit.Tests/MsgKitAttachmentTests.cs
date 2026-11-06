/////////////////////////////////////////////////////////////////////////////
// <copyright file="MsgKitAttachmentTests.cs" company="James John McGuire">
// Copyright © 2021 - 2026 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

#nullable enable

namespace DigitalZenWorks.Email.ToolKit.Tests;

using System.IO;
using MsgKit;
using NUnit.Framework;

/// <summary>
/// Checks data preservation for a valid inline attachment in upstream MsgKit.
/// These tests use its public API without activating Outlook.
/// </summary>
internal sealed class MsgKitAttachmentTests
{
	/// <summary>
	/// Verifies a valid inline attachment retains its content and metadata.
	/// </summary>
	/// <param name="contentId">The valid inline content ID.</param>
	[TestCase("legacy@example.test")]
	public void InlineAttachmentPreservesContentAndMetadata(string contentId)
	{
		byte[] expectedContent = [0x01, 0x02, 0x03, 0xff];
		using MemoryStream content = new(expectedContent);
		Attachments attachments = new();

		attachments.Add(
			content, "legacy.bin", isInline: true, contentId: contentId);

		Assert.That(attachments, Has.Count.EqualTo(1));
		Attachment attachment = attachments[0];
		Assert.That(attachment.FileName, Is.EqualTo("legacy.bin"));
		Assert.That(attachment.IsInline, Is.True);
		Assert.That(attachment.ContentId, Is.EqualTo(contentId));

		attachment.Stream.Position = 0;
		using MemoryStream copiedContent = new();
		attachment.Stream.CopyTo(copiedContent);
		byte[] actualContent = copiedContent.ToArray();
		Assert.That(actualContent, Is.EqualTo(expectedContent));
	}
}
