/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxMessageIndex.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx message indx class.
	/// </summary>
	public class DbxMessageIndex
	{
		/// <summary>
		/// Gets or sets the account associated with the message.
		/// </summary>
		/// <value>The account associated with the message.</value>
		public int Account { get; set; }

		/// <summary>
		/// Gets or sets the answered to message id.
		/// </summary>
		/// <value>The answered to message id.</value>
		public int AnswerId { get; set; }

		/// <summary>
		/// Gets or sets the body of the message.
		/// </summary>
		/// <value>The body of the message.</value>
		public string Body { get; set; }

		/// <summary>
		/// Gets or sets the pointer to the corresponding message.
		/// </summary>
		/// <value>The pointer to the corresponding message.</value>
		public string CorrespoindingMessage { get; set; }

		/// <summary>
		/// Gets or sets the flags of the message.
		/// </summary>
		/// <value>The flags of the message.</value>
		public int Flags { get; set; }

		/// <summary>
		/// Gets or sets the index is used for the Hotmail Http email accounts.
		/// </summary>
		/// <value>The index is used for the Hotmail Http
		/// email accounts.</value>
		/// <remarks>
		/// The This index is used for the Hotmail Http email accounts and
		/// stores a message id ("MSG982493141.24"). I don't know if other
		/// Http email accounts are using this index too..
		/// </remarks>
		public int HotmailIndex { get; set; }

		/// <summary>
		/// Gets or sets the message id.
		/// </summary>
		/// <value>The message id.</value>
		public uint Id { get; set; }

		/// <summary>
		/// Gets or sets the index of the message.
		/// </summary>
		/// <value>The index of the message.</value>
		public int Index { get; set; }

		/// <summary>
		/// Gets or sets the number of lines in the body.
		/// </summary>
		/// <value>The number of lines in the body.</value>
		public int LineCount { get; set; }

		/// <summary>
		/// Gets or sets the created or send time of the message.
		/// </summary>
		/// <value>The created or send time of the message.</value>
		public DateTime MessageTime { get; set; }

		/// <summary>
		/// Gets or sets the original subject of the message.
		/// </summary>
		/// <value>The original subject of the message.</value>
		public string OriginalSubject { get; set; }

		/// <summary>
		/// Gets or sets the priority of the eMail.
		/// </summary>
		/// <value>The priority of the eMail(1 high, 3 normal, 5 low).</value>
		public int Priority { get; set; }

		/// <summary>
		/// Gets or sets the recipient name.
		/// </summary>
		/// <value>The recipient name.</value>
		public string ReceiptentName { get; set; }

		/// <summary>
		/// Gets or sets the recipient email address.
		/// </summary>
		/// <value>The recipient email address.</value>
		public string ReceiptentEmailAddress { get; set; }

		/// <summary>
		/// Gets or sets the time message created/received.
		/// </summary>
		/// <value>The time message created/received.</value>
		public DateTime ReceivedTime { get; set; }

		/// <summary>
		/// Gets or sets the registry key for mail or news account.
		/// </summary>
		/// <value>The registry key for mail or news account
		/// (like "00000008").</value>
		public int RegistryKey { get; set; }

		/// <summary>
		/// Gets or sets the time message saved in this folder.
		/// </summary>
		/// <value>The time message saved in this folder.</value>
		public DateTime SavedInFolderTime { get; set; }

		/// <summary>
		/// Gets or sets the sender mail address and name.
		/// </summary>
		/// <value>The sender mail address and name.</value>
		public string Sender { get; set; }

		/// <summary>
		/// Gets or sets the sender mail address and name.
		/// </summary>
		/// <value>The sender mail address and name.</value>
		public string SenderEmailAddress { get; set; }

		/// <summary>
		/// Gets or sets the sender name.
		/// </summary>
		/// <value>The sender name.</value>
		public string SenderName { get; set; }

		/// <summary>
		/// Gets or sets the subject of the message.
		/// </summary>
		/// <value>The subject of the message.</value>
		public string Subject { get; set; }
	}
}
