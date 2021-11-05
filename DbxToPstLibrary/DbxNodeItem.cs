/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxNodeItem.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx node item class.
	/// </summary>
	public class DbxNodeItem
	{
		/// <summary>
		/// Gets or sets the node children index.
		/// </summary>
		/// <value>The node children index.</value>
		public int NodeChildrenIndex { get; set; }

		/// <summary>
		/// Gets or sets the node children count.
		/// </summary>
		/// <value>The node children count.</value>
		public int NodeChildrenCount { get; set; }

		/// <summary>
		/// Gets or sets the node value.
		/// </summary>
		/// <value>The node value.</value>
		public int NodeValue { get; set; }
	}
}
