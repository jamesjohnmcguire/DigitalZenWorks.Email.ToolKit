/////////////////////////////////////////////////////////////////////////////
// <copyright file="PstOutlook.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Microsoft.Office.Interop.Outlook;
using System;
using System.Runtime.InteropServices;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Provides support for interating with an Outlook PST file.
	/// </summary>
	public class PstOutlook
	{
		/// <summary>
		/// Create a new pst storage file.
		/// </summary>
		/// <param name="path">The path to the pst file.</param>
		/// <returns>A store object.</returns>
		public Store CreateStore(string path)
		{
			Application outlookApplication = new ();

			Store newPst = null;

			NameSpace outlookNamespace = outlookApplication.GetNamespace("mapi");

			outlookNamespace.Session.AddStore(path);

			foreach (Store store in outlookNamespace.Session.Stores)
			{
				if (store.FilePath == path)
				{
					newPst = store;
					break;
				}
			}

			return newPst;
		}
	}
}
