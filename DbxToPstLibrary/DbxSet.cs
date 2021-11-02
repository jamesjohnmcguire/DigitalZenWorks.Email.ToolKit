/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxSet.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	public class DbxSet
	{
		public DbxSet(string path)
		{
			byte[] fileBytes = File.ReadAllBytes(path);

			byte[] headerBytes = new byte[0x24bc];
			Array.Copy(fileBytes, headerBytes, 0x24bc);

			DbxHeader header = new DbxHeader(headerBytes);
		}
	}
}
