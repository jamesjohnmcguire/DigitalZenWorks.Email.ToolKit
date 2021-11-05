/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxException.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Runtime.Serialization;

namespace DbxToPstLibrary
{
	/// <summary>
	/// The dbx exception class.
	/// </summary>
	[Serializable]
	public class DbxException : Exception
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="DbxException"/> class.
		/// </summary>
		public DbxException()
		{ }

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxException"/> class.
		/// </summary>
		/// <param name="message">The message to include.</param>
		public DbxException(string message)
			: base(message)
		{ }

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxException"/> class.
		/// </summary>
		/// <param name="message">The message to include.</param>
		/// <param name="innerException">The inner exception.</param>
		public DbxException(string message, Exception innerException)
			: base(message, innerException)
		{ }

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxException"/> class.
		/// </summary>
		/// <param name="serializationInfo">The serialization info.</param>
		/// <param name="streamingContext">The streaming context.</param>
		protected DbxException(
			SerializationInfo serializationInfo,
			StreamingContext streamingContext)
		{
			throw new NotImplementedException();
		}
	}
}
