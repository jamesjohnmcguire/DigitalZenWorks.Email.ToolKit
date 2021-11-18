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
		private const int ItemsBase = 6;
		private const int NodeBaseAddressIndex = 0;
		private const int NodeIemCountIndex = 0x11;
		private const int TreeNodeSize = 0x27c;

		private static readonly ILog Log = LogManager.GetLogger(
			System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

		private readonly IList<uint> folderInformationIndexes =
			new List<uint>();

		/// <summary>
		/// Initializes a new instance of the <see cref="DbxTree"/> class.
		/// </summary>
		/// <param name="fileBytes">The bytes of the file.</param>
		/// <param name="rootNodeAddress">The address of the root node.</param>
		public DbxTree(byte[] fileBytes, uint rootNodeAddress)
		{
			ReadTree(fileBytes, rootNodeAddress);
		}

		/// <summary>
		/// Gets the folder information indexes list.
		/// </summary>
		/// <value>The folder information indexes list.</value>
		public IList<uint> FolderInformationIndexes
			{ get { return folderInformationIndexes; } }

		/// <summary>
		/// Reads the given bytes into a tree structure.
		/// </summary>
		/// <param name="fileBytes">The bytes of the file.</param>
		/// <param name="rootNodeAddress">The address of the root node.</param>
		public void ReadTree(byte[] fileBytes, uint rootNodeAddress)
		{
			if (fileBytes != null && rootNodeAddress != 0)
			{
				byte[] treeBytes = new byte[TreeNodeSize];

				Array.Copy(
					fileBytes, rootNodeAddress, treeBytes, 0, TreeNodeSize);

				// It will be easier to work with integers as opposed to bytes.
				int size = treeBytes.Length / sizeof(uint);
				uint[] treeArray = new uint[size];
				Buffer.BlockCopy(
					treeBytes, 0, treeArray, 0, treeBytes.Length);

				if (treeArray[0] != rootNodeAddress)
				{
					throw new DbxException("Wrong object marker!");
				}

				DbxTreeNode root = new ();
				root.NodeFileIndex = treeArray[NodeBaseAddressIndex];
				root.ChildrenNodesIndex = treeArray[2];

				// for root, should be 0
				root.ParentNodeIndex = treeArray[3];

				// recurse into sub tree.
				ReadTree(fileBytes, root.ChildrenNodesIndex);

				root.ItemCount = treeBytes[NodeIemCountIndex];

				for (int index = 0; index < root.ItemCount; index++)
				{
					DbxNodeItem item = SetIndexedValue(index, treeArray);

					// recurse into sub tree.
					ReadTree(fileBytes, item.NodeChildrenIndex);
				}
			}
		}

		/// <summary>
		/// Set indexed value method.
		/// </summary>
		/// <param name="index">The index of the item.</param>
		/// <param name="treeArray">The tree array of values.</param>
		/// <returns>A DbxNodeItem.</returns>
		public virtual DbxNodeItem SetIndexedValue(int index, uint[] treeArray)
		{
			// Each of the items occupy 3 ints (12 bytes) each,
			// starting from the 6th element.
			int offset = index * 3;
			offset += ItemsBase;

			DbxNodeItem item = new ();

			if (treeArray == null)
			{
				Log.Warn("tree arry is null");
			}
			else
			{
				item.NodeValue = treeArray[offset];

				if (item.NodeValue == 0)
				{
					Log.Warn("item node value is 0");
				}

				// also, add this to our list
				folderInformationIndexes.Add(item.NodeValue);

				offset++;
				item.NodeChildrenIndex = treeArray[offset];
			}

			return item;
		}
	}
}
