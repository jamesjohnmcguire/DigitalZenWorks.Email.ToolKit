/////////////////////////////////////////////////////////////////////////////
// <copyright file="CommandOption.cs" company="James John McGuire">
// Copyright © 2021 - 2022 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolKit.Application
{
	/// <summary>
	/// Represents a command line command option.
	/// </summary>
	public class CommandOption
	{
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
