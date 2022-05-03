/////////////////////////////////////////////////////////////////////////////
// <copyright file="HtmlEmail.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
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
			string pattern = @"(\\r\\n)+(?=\\r\\n</BODY>\\r\\n</HTML>$)";

			htmlBody = Regex.Replace(
				htmlBody, pattern, string.Empty, RegexOptions.ExplicitCapture);

			pattern = @"(<BR>\\r\\n)+(?=<BR>\\r\\n</FONT>\\r\\n" +
				@"</P>\\r\\n</BODY>\\r\\n</HTML>$)";

			htmlBody = Regex.Replace(
				htmlBody, pattern, string.Empty, RegexOptions.ExplicitCapture);

			return htmlBody;
		}
	}
}
