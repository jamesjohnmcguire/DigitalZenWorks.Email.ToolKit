/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandOption.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;

namespace DigitalZenWorks.Email.ToolKit.Application
{
	/// <summary>
	/// Represents a command line command option.
	/// </summary>
	public class CommandOption
	{
		/// <summary>
		/// Initializes a new instance of the <see cref="CommandOption"/>
		/// class.
		/// </summary>
		public CommandOption()
		{
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="CommandOption"/>
		/// class.
		/// </summary>
		/// <param name="shortName">The command short name.</param>
		/// <param name="longName">The command long name.</param>
		public CommandOption(string shortName, string longName)
		{
			ShortName = shortName;
			LongName = longName;
		}

		/// <summary>
		/// Gets or sets the long name.
		/// </summary>
		/// <value>The long name.</value>
		public string LongName { get; set; }

		/// <summary>
		/// Gets or sets the short name.
		/// </summary>
		/// <value>The short name.</value>
		public string ShortName { get; set; }
	}
}
