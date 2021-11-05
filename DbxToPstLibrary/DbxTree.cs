/////////////////////////////////////////////////////////////////////////////
// <copyright file="DbxTree.cs" company="James John McGuire">
// Copyright © 2021 James John McGuire. All Rights Reserved.
// </copyright>
/////////////////////////////////////////////////////////////////////////////

using Common.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbxToPstLibrary
{
	/// <summary>
	/// Dbx tree class.
	/// </summary>
	public class DbxTree
	{
		private const int NodeBaseAddressIndex = 0;
		private const int TreeNodeSize = 0x27c;

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxTree"/> class.
		/// </summary>
		/// <param name="treeBytes">The bytes of the tree in file.</param>
		/// <param name="rootNodeAddress">The address of the root node.</param>
		/// <param name="nodeCount">The count of nodes.</param>
		public DbxTree(byte[] treeBytes, int rootNodeAddress, int nodeCount)
		{
			if (treeBytes != null)
			{
				// It will be easier to work with integers as opposed to bytes.
				int size = treeBytes.Length / sizeof(int);
				int[] treeArray = new int[size];
				Buffer.BlockCopy(
					treeBytes, 0, treeArray, 0, treeBytes.Length);

				if (treeBytes[NodeBaseAddressIndex] != rootNodeAddress)
				{
					throw new DbxException("Wrong object marker!");
				}

			}
		}
	}
}
