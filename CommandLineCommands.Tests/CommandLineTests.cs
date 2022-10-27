/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandLineTests.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using DigitalZenWorks.CommandLine.Commands;
using NUnit.Framework;
using System;
using System.Collections.Generic;

[assembly: CLSCompliant(true)]

namespace CommandLineCommands.Tests
{
	/// <summary>
	/// Test class.
	/// </summary>
	public class CommandLineTests
	{
		private IList<Command> commands;

		/// <summary>
		/// One time set up method.
		/// </summary>
		[OneTimeSetUp]
		public void OneTimeSetUp()
		{
			commands = new List<Command>();

			Command help = new ("help");
			help.Description = "Show this information";
			commands.Add(help);

			CommandOption option = new ("e", "encoding");
			IList<CommandOption> options = new List<CommandOption>();
			options.Add(option);

			Command commandOne = new (
				"command-one",
				options,
				1,
				"A command with an option that has an option.");
			commands.Add(commandOne);

			Command commandTwo = new (
				"command-two",
				null,
				1,
				"A command with no options");
			commands.Add(commandTwo);

			CommandOption dryRun = new ("n", "dryrun");
			options = new List<CommandOption>();
			options.Add(dryRun);

			Command commandThree = new (
				"command-three",
				options,
				0,
				"A command with no parameters");
			commands.Add(commandThree);

			Command commandFour = new (
				"command-four",
				null,
				2,
				"A command with 2 minimum required parameters");
			commands.Add(commandFour);

			Command commandFive = new (
				"command-five",
				null,
				4,
				"A command with 4 minimum required parameters");
			commands.Add(commandFive);

			CommandOption flush = new ("s", "flush");
			options = new List<CommandOption>();
			options.Add(dryRun);
			options.Add(flush);

			Command commandSix = new (
				"command-six",
				options,
				0,
				"A command with multiple options");
			commands.Add(commandSix);

			CommandOption encoding = new ("e", "encoding", true);
			options = new List<CommandOption>();
			options.Add(encoding);

			Command commandSeven = new (
				"command-seven",
				options,
				1,
				"A command with an option that has a value.");
			commands.Add(commandSeven);
		}

		/// <summary>
		/// One time tear down method.
		/// </summary>
		[OneTimeTearDown]
		public void OneTimeTearDown()
		{
		}

		/// <summary>
		/// Set up method.
		/// </summary>
		[SetUp]
		public void Setup()
		{
		}

		/// <summary>
		/// Sanity test.
		/// </summary>
		[Test]
		public void SanityTest()
		{
			Assert.Pass();
		}

		/// <summary>
		/// Help test.
		/// </summary>
		[Test]
		public void HelpTest()
		{
			string[] arguments = { "help" };

			CommandLineArguments commandLine = new (commands, arguments);

			Assert.True(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.NotNull(command);

			Assert.AreEqual("help", command.Name);
		}

		/// <summary>
		/// Option simple no option test.
		/// </summary>
		[Test]
		public void OptionSimpleNoOptionTest()
		{
			string[] arguments = { "command-one", "%USERPROFILE%" };

			CommandLineArguments commandLine = new (commands, arguments);

			Assert.True(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.NotNull(command);

			Assert.AreEqual("command-one", command.Name);

			IList<CommandOption> options = command.Options;

			Assert.AreEqual(options.Count, 0);
		}

		/// <summary>
		/// Option simple fail no parameter test.
		/// </summary>
		[Test]
		public void OptionSimpleFailNoParameterTest()
		{
			string[] arguments = { "command-one", "-e" };

			CommandLineArguments commandLine = new (commands, arguments);

			Assert.False(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.Null(command);
		}

		/// <summary>
		/// Option simple short option first test.
		/// </summary>
		[Test]
		public void OptionSimpleShortOptionFirstTest()
		{
			string[] arguments = { "command-one", "-e", "%USERPROFILE%" };

			CommandLineArguments commandLine = new (commands, arguments);

			Assert.True(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.NotNull(command);

			Assert.AreEqual("command-one", command.Name);

			IList<CommandOption> options = command.Options;

			Assert.Greater(options.Count, 0);
		}

		/// <summary>
		/// Option simple short option last test.
		/// </summary>
		[Test]
		public void OptionSimpleShortOptionLastTest()
		{
			string[] arguments = { "command-one", "%USERPROFILE%", "-e" };

			CommandLineArguments commandLine = new (commands, arguments);

			Assert.True(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.NotNull(command);

			Assert.AreEqual("command-one", command.Name);

			IList<CommandOption> options = command.Options;

			Assert.Greater(options.Count, 0);
		}

		/// <summary>
		/// Option simple long option first test.
		/// </summary>
		[Test]
		public void OptionSimpleLongOptionFirstTest()
		{
			string[] arguments =
				{ "command-one", "--encoding", "%USERPROFILE%" };

			CommandLineArguments commandLine = new(commands, arguments);

			Assert.True(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.NotNull(command);

			Assert.AreEqual("command-one", command.Name);

			IList<CommandOption> options = command.Options;

			Assert.Greater(options.Count, 0);
		}

		/// <summary>
		/// Option simple long option last test.
		/// </summary>
		[Test]
		public void OptionSimpleLongOptionLastTest()
		{
			string[] arguments =
				{ "command-one", "%USERPROFILE%", "--encoding" };

			CommandLineArguments commandLine = new (commands, arguments);

			Assert.True(commandLine.ValidArguments);

			Command command = commandLine.Command;
			Assert.NotNull(command);

			Assert.AreEqual("command-one", command.Name);

			IList<CommandOption> options = command.Options;

			Assert.Greater(options.Count, 0);
		}
	}
}
