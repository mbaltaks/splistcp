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
	class MBSPListItemMap /*: IComparable*/
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
						string loginName = MBSPSiteMap.GetLoginNameFromSharePointID(value.ToString(), sourceSiteURL, Program.getSourceCredentials());
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
				sourceListName, sourceListNameURL, sourceFolderPath, Program.getSourceCredentials());

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
			// Sorting doesn't solve this, too many gaps in source IDs.
			//subItems.Sort();
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
				listService.Credentials = Program.getSourceCredentials();
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


		// This assumes Aus/UK date format - need a way to determine site regional setting.
		// There is no way, we'll have to guess. What a hopeless API.
		public string SPAttributeDateFormat(string date)
		{
			string d = date;
			string day = d.Substring(0, d.IndexOf('/'));
			d = d.Substring(d.IndexOf('/') + 1);
			string month = d.Substring(0, d.IndexOf('/'));
			int monthnum = System.Convert.ToInt32(month, 10);
			if (MBSPSiteMap.versionsUseUSDates || (monthnum > 12))
			{
				string s = day;
				day = month;
				month = s;
				MBSPSiteMap.versionsUseUSDates = true;
			}
			d = d.Substring(d.IndexOf('/') + 1);
			string year = d.Substring(0, d.IndexOf(' '));
			d = d.Substring(d.IndexOf(' ') + 1);
			string hour = d.Substring(0, d.IndexOf(':'));
			d = d.Substring(d.IndexOf(':') + 1);
			string minute = d.Substring(0, d.IndexOf(' '));
			d = d.Substring(d.IndexOf(' ') + 1);
			int h = System.Convert.ToInt32(hour, 10);
			if (d.Equals("PM"))
			{
				if (h < 12)
				{
					h += 12;
				}
				hour = h.ToString();
			}
			else if (d.Equals("AM"))
			{
				if (h.Equals(12))
				{
					h = 0;
				}
				hour = h.ToString();
			}
			string newDate = year + "-" + month + "-" + day + " " + hour + ":" + minute + ":00";
			return newDate;
		}


		public void GetListItemVersions(string sourceSiteURL, string listName, 
			string folderPath, string fileName, string destFolderPath, ArrayList attributeNames,
			Hashtable attributes)
		{
			string[] paths = new string[] { listName, folderPath, fileName };
			string filePath = MBSPListMap.CombinePaths(paths);
			SharePointVersionsWebService.Versions v = new SharePointVersionsWebService.Versions();
			v.Url = sourceSiteURL + "/_vti_bin/Versions.asmx";
			v.Credentials = Program.getSourceCredentials();
			XmlNode versions = v.GetVersions(filePath);
			string fileURL = "";
			string editor = attributes["ows_Editor"].ToString();
			string modified = attributes["ows_Modified"].ToString();
			foreach (XmlNode child in versions)
			{
				if (child.Name.Equals("result"))
				{
					if (!child.Attributes["version"].Value.ToString().StartsWith("@"))
					{
						fileURL = child.Attributes["url"].Value.ToString();
						SPListItem newItem = AddRemoteFile(fileName, fileURL, destFolderPath);
						// now set created date and created by for this version
						attributes["ows_Editor"] = MBSPSiteMap.GetSharePointIDFromLoginName(child.Attributes["createdBy"].Value.ToString(), sourceSiteURL, Program.getSourceCredentials());
						attributes["ows_Modified"] = SPAttributeDateFormat(child.Attributes["created"].Value.ToString());
						SetListItemAttributes(attributeNames, attributes, newItem, sourceSiteURL);
					}
				}
			}
			attributes["ows_Editor"] = editor;
			attributes["ows_Modified"] = modified;
		}


		public SPListItem AddRemoteFile(string itemName, string fileURL, string destFolderPath)
		{
			System.Net.WebClient client = new System.Net.WebClient();
			client.Credentials = Program.getSourceCredentials();
			string localPath = Program.tempFilePath + "/" + itemName;
			System.Console.Out.WriteLine("");
			System.Console.Out.WriteLine("Downloading " + fileURL);
			client.DownloadFile(fileURL, localPath);

			FileStream localFile = File.OpenRead(localPath);
			//metadataTable.Add("vti_title", title);
			SPFolder f = listMap.EnsureFolderPathExists(listMap.destList.RootFolder, destFolderPath);
			System.Console.Out.WriteLine("Adding " + f.ServerRelativeUrl + "/"
				+ itemName);
			Hashtable metadataTable = new Hashtable();
			SPFileCollection files = f.Files;
			SPFile newFile = files.Add(f.ServerRelativeUrl + "/" + itemName, localFile, metadataTable, true);
			localFile.Close();
			File.Delete(localPath);
			return newFile.Item;
		}


		public void SetListItemAttributes(ArrayList attributeNames, Hashtable attributes,
			SPListItem newItem, string sourceSiteURL)
		{
			bool versions = newItem.ParentList.EnableVersioning;
			newItem.ParentList.EnableVersioning = false;
			newItem.ParentList.Update();
			foreach (Object attributeName in attributeNames)
			{
				SetListItemAttribute(newItem, attributeName.ToString(),
					attributes[attributeName], listMap.listFields, listMap.newListFields,
					sourceSiteURL);
			}
			/*try
			{*/
			newItem.Update();
			newItem.ParentList.EnableVersioning = versions;
			newItem.ParentList.Update();
			/*}
			catch (Exception e)
			{
				Console.WriteLine("** There was a problem writing attributes for this item:");
				Console.WriteLine(e.Message);
			}*/
		}


		public void CopyData(string sourceSiteURL, string sourceListNameURL)
		{
			System.Net.WebClient client = new System.Net.WebClient();
			client.Credentials = Program.getSourceCredentials();
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
						if (!Program.skipOldVersions)
						{
							GetListItemVersions(sourceSiteURL, sourceListNameURL,
								subItem.sourceFolderPath, subItem.itemName, subItem.destFolderPath,
								subItem.attributeNames, subItem.attributes);
						}
						try
						{
							newItem = AddRemoteFile(subItem.itemName, fileURL, subItem.destFolderPath);
						}
						catch (Exception e)
						{
							Console.WriteLine(e.Message);
							return;
						}
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
						SetListItemAttributes(subItem.attributeNames, subItem.attributes, newItem, sourceSiteURL);
					}
				}
			}
		}


		/*
		public int CompareTo(object obj)
		{
			MBSPListItemMap listItem = (MBSPListItemMap) obj;
			int myID = System.Convert.ToInt32(this.attributes["ows_ID"].ToString(), 10);
			int theirID = System.Convert.ToInt32(listItem.attributes["ows_ID"].ToString());
			return myID.CompareTo(theirID);
		}
		*/
	}
}
