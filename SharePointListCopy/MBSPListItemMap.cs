/*
 * Copyright © 2008 Michael Baltaks
 *
 * License:
 * This file is part of splistcp.
 *
 * splistcp is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 * 
 * splistcp is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with splistcp; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301 USA
 *
 */

using System;
using System.Collections;
using System.IO;
using System.Xml;
using Microsoft.SharePoint;


namespace SharePointListCopy
{
	class MBSPListItemMap
	{
		string itemName = "";
		string sourceFolderPath = "";
		string destFolderPath = "";
		bool hasSubItems = false;
		ArrayList subItems;
		bool hasFile = false;
		bool hasAttachments = false;
		ArrayList attachmentURLs = new ArrayList();
		ArrayList attributeNames;
		Hashtable attributes;
		MBSPListMap listMap;


		public MBSPListItemMap(MBSPListMap aListMap, string aDestFolderPath) : this(aListMap, aDestFolderPath, false)
		{
		}


		public MBSPListItemMap(MBSPListMap aListMap, string aDestFolderPath, bool aHasSubItems)
		{
			subItems = new ArrayList();
			attributeNames = new ArrayList();
			attributes = new Hashtable();
			listMap = aListMap;
			destFolderPath = aDestFolderPath;
			hasSubItems = aHasSubItems;
		}


		bool SetListItemAttribute(SPListItem item, string name, Object value,
			Hashtable listFields, Hashtable newListFields, string sourceSiteURL)
		{
			if (name.StartsWith("ows_"))
			{
				name = name.Substring(4);
			}
			string displayName = listFields[name].ToString();

			// Only add fields in newListFields
			bool cached = newListFields.ContainsKey(displayName);
			if (displayName.Equals("Type") && name.Equals("DocIcon"))
			{
				cached = false;
			}
			if (cached)
			{
				string newInternalName = newListFields[displayName].ToString();
				if (newInternalName.Length > 0)
				{
					SPField f = item.Fields.GetField(newInternalName);
					if (f.Type.Equals(SPFieldType.User))
					{
						string loginName = MBSPSiteMap.GetLoginNameFromSharePointID(value.ToString(), sourceSiteURL);
						string fullName = MBSPSiteMap.GetFullNameFromSharePointID(value.ToString());
						value = MBSPSiteMap.EnsureAUserExists(loginName, fullName, item.ParentList.ParentWeb);
					}
					item[newInternalName] = listMap.ReplaceValues(value.ToString());
					return true;
				}
			}
			return false;
		}


		// Look for all the items in this part of the list, and add them
		// to the subItems ArrayList.
		public void GetAllSubItems(SharePointListsWebService.Lists listService,
			string sourceListName, string sourceListNameURL)
		{
			XmlNode listNode = listMap.GetListItems(itemName,
				sourceListName, sourceListNameURL, sourceFolderPath);

			// Xml handling from http://blog.andyjohnson.org/?page_id=34
			String xpq = "//*[@*]"; //get all nodes
			XmlNodeList allNodes = listNode.SelectNodes(xpq);

			for (int i = 1; i < allNodes.Count; i++) // first node is whitespace
			{
				XmlNode listItemNode = allNodes[i];
				if (listItemNode.Attributes != null)
				{
					MBSPListItemMap newItem = GetListItem(listItemNode, listService, sourceListName);
					subItems.Add(newItem);
				}
			}

			foreach (MBSPListItemMap item in subItems)
			{
				if (item.hasSubItems)
				{
					item.GetAllSubItems(listService, sourceListName, sourceListNameURL);
				}
			}
		}


		public MBSPListItemMap GetListItem(XmlNode node, SharePointListsWebService.Lists listService,
			string sourceListName)
		{
			string objectType = "";
			MBSPListItemMap newItem = new MBSPListItemMap(listMap, "");

			foreach (XmlAttribute attr in node.Attributes)
			{
				newItem.attributes.Add(attr.Name, attr.Value);
				newItem.attributeNames.Add(attr.Name);
				switch (attr.Name)
				{
					case "ows_LinkFilename":
						newItem.itemName = attr.Value;
						break;
					case "ows_FSObjType":
						objectType = attr.Value;
						if (objectType.EndsWith("1"))
						{
							newItem.hasSubItems = true;
						}
						else if (objectType.EndsWith("0"))
						{
							newItem.hasFile = true;
						}
						break;
					case "ows_Attachments":
						if (!attr.Value.ToString().Equals("0"))
						{
							newItem.hasAttachments = true;
						}
						break;
				}
				//System.Console.Out.WriteLine(attr.Name + ": " + attr.Value);
			}
			if (newItem.hasAttachments)
			{
				XmlNode attachmentsNode = listService.GetAttachmentCollection(sourceListName,
					newItem.attributes["ows_ID"].ToString());
				foreach (XmlNode att in attachmentsNode)
				{
					newItem.attachmentURLs.Add(att.FirstChild.Value.ToString());
				}
			}
			string[] sourcePaths = new string[] { sourceFolderPath, itemName };
			string[] destPaths = new string[] { destFolderPath, itemName };
			newItem.sourceFolderPath = MBSPListMap.CombinePaths(sourcePaths);
			newItem.destFolderPath = MBSPListMap.CombinePaths(destPaths);
			return newItem;
		}


		public void CopyData(string sourceSiteURL, string sourceListNameURL)
		{
			System.Net.WebClient client = new System.Net.WebClient();
			client.Credentials = System.Net.CredentialCache.DefaultCredentials;
			System.IO.Directory.CreateDirectory(Program.tempFilePath);
			foreach (MBSPListItemMap subItem in subItems)
			{
				SPListItem newItem = null;
				if (Program.avoidDuplicates)
				{
					foreach (SPListItem item in listMap.destList.Items)
					{
						string internalName = listMap.newListFields[listMap.listKeyField].ToString();
						string sourceInternalName = listMap.reverseListFields[listMap.listKeyField].ToString();
						object o = item[internalName];
						string existingKey = o.ToString();
						string newKey = subItem.attributes["ows_" + sourceInternalName].ToString();
						if (existingKey.Equals(newKey))
						{
							newItem = item;
							Console.WriteLine("Item " + item.DisplayName + " already exists");
							break;
						}
					}
				}
				if (newItem == null)
				{
					if (subItem.hasFile)
					{
						string[] paths = new string[] { sourceSiteURL, 
						sourceListNameURL, subItem.sourceFolderPath, subItem.itemName };
						string fileURL = MBSPListMap.CombinePaths(paths);
						string localPath = Program.tempFilePath + "/" + subItem.itemName;
						System.Console.Out.WriteLine("");
						System.Console.Out.WriteLine("Downloading " + fileURL);
						try
						{
							client.DownloadFile(fileURL, localPath);
						}
						catch (Exception e)
						{
							Console.WriteLine(e.Message);
							return;
						}

						FileStream localFile = File.OpenRead(localPath);
						//metadataTable.Add("vti_title", title);
						SPFolder f = listMap.EnsureFolderPathExists(listMap.destList.RootFolder, subItem.destFolderPath);
						System.Console.Out.WriteLine("Adding " + f.ServerRelativeUrl + "/"
							+ subItem.itemName);
						Hashtable metadataTable = new Hashtable();
						SPFileCollection files = f.Files;
						SPFile newFile = files.Add(f.ServerRelativeUrl + "/" + subItem.itemName, localFile, metadataTable, true);
						localFile.Close();
						File.Delete(localPath);
						newItem = newFile.Item;
					}
					else if (subItem.hasSubItems)
					{
						string[] paths = new string[] { subItem.destFolderPath, subItem.itemName };
						SPFolder f = listMap.EnsureFolderPathExists(listMap.destList.RootFolder, MBSPListMap.CombinePaths(paths));
						newItem = f.Item;
						subItem.CopyData(sourceSiteURL, sourceListNameURL);
					}
					else
					{
						newItem = listMap.destList.Items.Add();
					}
					if (subItem.hasAttachments && newItem != null)
					{
						foreach (string downloadUrl in subItem.attachmentURLs)
						{
							string fileName = downloadUrl.Substring(downloadUrl.LastIndexOf('/') + 1);
							string downloadPath = Program.tempFilePath + "/" + fileName;
							System.Console.Out.WriteLine("");
							System.Console.Out.WriteLine("Downloading " + downloadUrl);
							try
							{
								client.DownloadFile(downloadUrl, downloadPath);
							}
							catch (Exception e)
							{
								Console.WriteLine(e.Message);
							}
							Console.WriteLine("Attaching " + fileName);
							byte[] fileContents = MBSPSiteMap.ByteArrayFromFilePath(downloadPath);
							newItem.Attachments.Add(fileName, fileContents);
							File.Delete(downloadPath);
						}
					}
					if (newItem != null)
					{
						foreach (Object attributeName in subItem.attributeNames)
						{
							SetListItemAttribute(newItem, attributeName.ToString(),
								subItem.attributes[attributeName], listMap.listFields, listMap.newListFields,
								sourceSiteURL);
						}
						/*try
						{*/
						newItem.Update();
						/*}
						catch (Exception e)
						{
							Console.WriteLine("** There was a problem writing attributes for this item:");
							Console.WriteLine(e.Message);
						}*/
					}
				}
			}
		}
	}
}
