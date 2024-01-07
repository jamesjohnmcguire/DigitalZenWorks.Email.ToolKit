/////////////////////////////////////////////////////////////////////////////
// <copyright file="HtmlEmail.cs" company="James John McGuire">
// Copyright © 2021 - 2024 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DigitalZenWorks.Email.ToolKit
{
	/// <summary>
	/// Represents a HTML email body.
	/// </summary>
	public static class HtmlEmail
	{
		/// <summary>
		/// Trim the end of the HTML body.
		/// </summary>
		/// <param name="htmlBody">The HTML body to trim.</param>
		/// <returns>The trimmed HTML body.</returns>
		public static string Trim(string htmlBody)
		{
			string pattern = @"(\r{0,1}\n)+(?=\r{0,1}\n<\/BODY>" +
				@"\r{0,1}\n<\/HTML>$)";

			htmlBody = Regex.Replace(
				htmlBody,
				pattern,
				string.Empty,
				RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase);

			pattern = @"(<BR>\r{0,1}\n)+(?=<BR>\r{0,1}\n<\/FONT>\r{0,1}\n" +
				@"<\/P>\r{0,1}\n<\/BODY>\r{0,1}\n<\/HTML>$)";

			htmlBody = Regex.Replace(
				htmlBody,
				pattern,
				string.Empty,
				RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase);

			return htmlBody;
		}
	}
}
