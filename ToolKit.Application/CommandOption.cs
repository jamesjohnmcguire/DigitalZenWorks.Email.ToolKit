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
		private bool requiresParameter;

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
		/// <param name="requiresParameter">A value indicating whether this
		/// option requires a parameter or not.</param>
		public CommandOption(
			string shortName, string longName, bool requiresParameter = false)
		{
			ShortName = shortName;
			LongName = longName;

			this.requiresParameter = requiresParameter;
		}

		/// <summary>
		/// Gets or sets the option's argument index.
		/// </summary>
		/// <value>The option's argument index.</value>
		public int ArgumentIndex { get; set; }

		/// <summary>
		/// Gets or sets the long name.
		/// </summary>
		/// <value>The long name.</value>
		public string LongName { get; set; }

		/// <summary>
		/// Gets or sets the option parameter.
		/// </summary>
		/// <value>The option parameter.</value>
		public string Parameter { get; set; }

		/// <summary>
		/// Gets a value indicating whether a value indicating whether this
		/// option requires a parameter or not.
		/// </summary>
		/// <value>A value indicating whether this option requires a
		/// parameter or not.</value>
		public bool RequiresParameter { get { return requiresParameter; } }

		/// <summary>
		/// Gets or sets the short name.
		/// </summary>
		/// <value>The short name.</value>
		public string ShortName { get; set; }
	}
}
