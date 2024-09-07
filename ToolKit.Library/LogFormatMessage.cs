/////////////////////////////////////////////////////////////////////////////
// <copyright file="LogFormatMessage.cs" company="James John McGuire">
// Copyright © 2021 - 2024 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System.Globalization;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Provides support for compiling a formatted message and
	/// logging that message.
	/// </summary>
	public static class LogFormatMessage
	{
		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		/// <summary>
		/// Sends an error log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		public static void Error(
			string template, string parameter1)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1);

			Log.Error(message);
		}

		/// <summary>
		/// Sends an error log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		public static void Error(
			string template, string parameter1, string parameter2)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2);

			Log.Error(message);
		}

		/// <summary>
		/// Sends an error log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		/// <param name="parameter3">The third parameter.</param>
		public static void Error(
			string template,
			string parameter1,
			string parameter2,
			string parameter3)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2,
				parameter3);

			Log.Error(message);
		}

		/// <summary>
		/// Sends an error log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		/// <param name="parameter3">The third parameter.</param>
		/// <param name="parameter4">The fourth parameter.</param>
		public static void Error(
			string template,
			string parameter1,
			string parameter2,
			string parameter3,
			string parameter4)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2,
				parameter3,
				parameter4);

			Log.Error(message);
		}

		/// <summary>
		/// Sends an info log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		public static void Info(
			string template, string parameter1, string parameter2)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2);

			Log.Info(message);
		}

		/// <summary>
		/// Sends an info log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		/// <param name="parameter3">The third parameter.</param>
		public static void Info(
			string template,
			string parameter1,
			string parameter2,
			string parameter3)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2,
				parameter3);

			Log.Info(message);
		}

		/// <summary>
		/// Sends an info log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		/// <param name="parameter3">The third parameter.</param>
		/// <param name="parameter4">The fourth parameter.</param>
		public static void Info(
			string template,
			string parameter1,
			string parameter2,
			string parameter3,
			string parameter4)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2,
				parameter3,
				parameter4);

			Log.Info(message);
		}

		/// <summary>
		/// Sends an info log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		/// <param name="parameter3">The third parameter.</param>
		/// <param name="parameter4">The fourth parameter.</param>
		/// <param name="parameter5">The fifth parameter.</param>
		public static void Info(
			string template,
			string parameter1,
			string parameter2,
			string parameter3,
			string parameter4,
			string parameter5)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2,
				parameter3,
				parameter4,
				parameter5);

			Log.Info(message);
		}

		/// <summary>
		/// Sends an error log message.
		/// </summary>
		/// <param name="template">The format template to use.</param>
		/// <param name="parameter1">The first parameter.</param>
		/// <param name="parameter2">The second parameter.</param>
		/// <param name="parameter3">The third parameter.</param>
		/// <param name="parameter4">The fourth parameter.</param>
		public static void Warn(
			string template,
			string parameter1,
			string parameter2,
			string parameter3,
			string parameter4)
		{
			string message = string.Format(
				CultureInfo.InvariantCulture,
				template,
				parameter1,
				parameter2,
				parameter3,
				parameter4);

			Log.Warn(message);
		}
	}
}
