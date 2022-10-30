/////////////////////////////////////////////////////////////////////////////
// <copyright file="Command.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;

namespace DigitalZenWorks.CommandLine.Commands
{
	/// <summary>
	/// Represents a command line command.
	/// </summary>
	public class Command
	{
		private readonly string name;
		private readonly IList<CommandOption> options;
		private readonly int parameterCount;

		private string description;
		private IList<string> parameters = new List<string>();

		/// <summary>
		/// Initializes a new instance of the <see cref="Command"/> class.
		/// </summary>
		/// <param name="name">The command name.</param>
		public Command(string name)
		{
			this.name = name;

			if (options == null)
			{
				options = new List<CommandOption>();
			}
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Command"/> class.
		/// </summary>
		/// <param name="name">The command name.</param>
		/// <param name="options">The command options.</param>
		/// <param name="parameters">The command parameters.</param>
		public Command(
			string name,
			IList<CommandOption> options,
			IList<string> parameters)
			: this(name)
		{
			this.options = options;
			this.parameters = parameters;
		}

		/// <summary>
		/// Initializes a new instance of the <see cref="Command"/> class.
		/// </summary>
		/// <param name="name">The command name.</param>
		/// <param name="options">The command options.</param>
		/// <param name="parameterCount">The command required parameter
		/// count.</param>
		/// <param name="description">The command description.</param>
		public Command(
			string name,
			IList<CommandOption> options,
			int parameterCount,
			string description)
			: this(name)
		{
			this.options = options;
			this.parameterCount = parameterCount;
			this.description = description;
		}

		/// <summary>
		/// Gets or sets the command description.
		/// </summary>
		/// <value>The command description.</value>
		public string Description
		{
			get { return description; }
			set { description = value; }
		}

		/// <summary>
		/// Gets the command name.
		/// </summary>
		/// <value>The command name.</value>
		public string Name { get { return name; } }

		/// <summary>
		/// Gets the command options.
		/// </summary>
		/// <value>The command options.</value>
		public IList<CommandOption> Options { get { return options; } }

		/// <summary>
		/// Gets the command parameter count.
		/// </summary>
		/// <value>The command parameter count.</value>
		public int ParameterCount { get { return parameterCount; } }

		/// <summary>
		/// Gets the command parameters.
		/// </summary>
		/// <value>The command parameters.</value>
		public IList<string> Parameters { get { return parameters; } }

		/// <summary>
		/// Does option exist.
		/// </summary>
		/// <param name="shortName">The short name to search for.</param>
		/// <param name="longName">The long name to search for.</param>
		/// <returns>A value indicating whether the option exists
		/// or not.</returns>
		public bool DoesOptionExist(string shortName, string longName)
		{
			bool optionExists = false;

			CommandOption optionFound = GetOption(shortName, longName);

			if (optionFound != null)
			{
				optionExists = true;
			}

			return optionExists;
		}

		/// <summary>
		/// Get option.
		/// </summary>
		/// <param name="shortName">The short name to search for.</param>
		/// <param name="longName">The long name to search for.</param>
		/// <returns>The found option, if it exists.</returns>
		public CommandOption GetOption(string shortName, string longName)
		{
			List<CommandOption> optionsList = options.ToList();

			CommandOption option = optionsList.Find(option =>
				(option.ShortName != null &&
				option.ShortName.Equals(
					shortName, StringComparison.Ordinal)) ||
				(option.LongName != null &&
				option.LongName.Equals(
					longName, StringComparison.Ordinal)));

			return option;
		}
	}
}
