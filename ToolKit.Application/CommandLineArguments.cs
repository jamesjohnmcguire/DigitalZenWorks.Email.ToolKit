/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandLineArguments.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;

namespace DigitalZenWorks.Email.ToolKit.Application
{
	/// <summary>
	/// Represents a set of command line arguments.
	/// </summary>
	public class CommandLineArguments
	{
		private readonly string[] arguments;
		private readonly IList<Command> commands;

		private string command;
		private string errorMessage;
		private string invalidOption;
		private bool validArguments;

		/// <summary>
		/// Initializes a new instance of the
		/// <see cref="CommandLineArguments"/> class.
		/// </summary>
		/// <param name="commands">A list of valid commands.</param>
		/// <param name="arguments">The array of command line
		/// arguments.</param>
		public CommandLineArguments(
			IList<Command> commands, string[] arguments)
		{
			this.commands = commands;
			this.arguments = arguments;

			validArguments = ValidateArguments();
		}

		/// <summary>
		/// Gets the error message, if any.
		/// </summary>
		/// <value>The error message, if any.</value>
		public string ErrorMessage { get { return errorMessage; } }

		/// <summary>
		/// Gets or sets the usage statement.
		/// </summary>
		/// <value>The usage statement.</value>
		public string UsageStatement { get; set; }

		/// <summary>
		/// Gets a value indicating whether a value indicating whether the
		/// arguments are valid or not.
		/// </summary>
		/// <value>A value indicating whether the arguments are valid
		/// or not.</value>
		public bool ValidArguments { get { return validArguments; } }

		private bool IsValidOption(
			Command command, CommandOption option)
		{
			bool isValid = false;

			foreach (CommandOption validOption in command.Options)
			{
				if (option.LongName.Equals(
					validOption.LongName, StringComparison.Ordinal) ||
					option.ShortName.Equals(
					validOption.ShortName, StringComparison.Ordinal))
				{
					isValid = true;
					break;
				}
			}

			if (isValid == false)
			{
				if (!string.IsNullOrWhiteSpace(option.LongName))
				{
					invalidOption = option.LongName;
				}
				else
				{
					invalidOption = option.ShortName;
				}
			}

			return isValid;
		}

		private IList<CommandOption> GetOptions(Command command)
		{
			IList<CommandOption> options = new List<CommandOption>();

			foreach (string argument in arguments)
			{
				if (argument.StartsWith('-'))
				{
					string optionName = argument.TrimStart('-');

					CommandOption option = new ();
					if (argument.StartsWith("--", StringComparison.Ordinal))
					{
						option.LongName = optionName;
					}
					else
					{
						option.ShortName = optionName;
					}

					options.Add(option);
				}
			}

			return options;
		}

		private IList<string> GetParameters(Command command)
		{
			IList<string> parameters = new List<string>();

			for (int index = 1; index < arguments.Length; index++)
			{
				string argument = arguments[index];

				if (!argument.StartsWith('-'))
				{
					parameters.Add(argument);
				}
			}

			return parameters;
		}

		private bool ValidateArguments()
		{
			bool areValid = false;
			bool isValidCommand = false;
			Command validatedCommand = null;

			command = arguments[0];

			foreach (Command validCommand in commands)
			{
				if (command.Equals(
					validCommand.Name, StringComparison.Ordinal))
				{
					validatedCommand = validCommand;
					isValidCommand = true;
					break;
				}
			}

			if (isValidCommand == false)
			{
				errorMessage = "Uknown command.";
			}
			else
			{
				bool areOptionsValid = true;

				IList<CommandOption> commandOptions =
					GetOptions(validatedCommand);

				foreach (CommandOption option in commandOptions)
				{
					bool isValid = IsValidOption(validatedCommand, option);

					if (isValid == false)
					{
						areOptionsValid = false;
						break;
					}
				}

				if (areOptionsValid == false)
				{
					errorMessage = string.Format(
						CultureInfo.InvariantCulture,
						"Uknown option:{0}.",
						invalidOption);
				}
				else
				{
					IList<string> parameters = GetParameters(validatedCommand);

					if (parameters.Count == validatedCommand.ParameterCount)
					{
						areValid = true;
					}
					else
					{
						errorMessage = "Incorrect amount of parameters.";
					}
				}
			}

			return areValid;
		}
	}
}
