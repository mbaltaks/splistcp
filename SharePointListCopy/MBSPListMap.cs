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
using System.Globalization;
using Microsoft.SharePoint;

namespace SharePointListCopy
{
	class MBSPListMap
	{
		string sourceSiteURL = "";
		string sourceListNameURL = "";
		string sourceListName = "";
		string destSiteURL = "";
		string destListName = "";
		string destFolderPath = "";
		SharePointListsWebService.Lists listService = new SharePointListsWebService.Lists();
		SPWeb web;
		public SPList destList;
		string sourceListDescription = "";
		string destListDescription = "";
		DateTime sourceListCreated = DateTime.Now;
		DateTime sourceListModified = DateTime.Now;
		DateTime destListCreated = DateTime.Now;
		DateTime destListModified = DateTime.Now;
		string sourceListAuthor = "";
		string destListAuthor = "";
		SPListTemplateType sourceListType = SPListTemplateType.DocumentLibrary;
		SPListTemplateType destListType = SPListTemplateType.DocumentLibrary;
		public Hashtable listFields = new Hashtable();
		public Hashtable reverseListFields = new Hashtable();
		public Hashtable newListFields = new Hashtable();
		public Hashtable reverseNewListFields = new Hashtable();
		bool newList = false;
		string listServiceURL = "/_vti_bin/Lists.asmx";
		public string listKeyField = "";
		string listKeyFieldInternal = "";
		bool wrongType = false;
		MBSPListItemMap topLevel;
		Hashtable replacements;


		public MBSPListMap(string raw_source, string raw_dest_site, string raw_dest_path,
			Hashtable aReplacements)
		{
			Console.WriteLine("");
			Console.WriteLine("Copying list " + raw_source);
			try
			{
				ParseURLs(raw_source, raw_dest_site, raw_dest_path);
			}
			catch
			{
				Program.printUsage();
				return;
			}
			listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			//listService.Credentials = new System.Net.NetworkCredential(username, password, domain);
			GetRealListName();
			web = GetSPWeb(destSiteURL);
			if (web == null)
			{
				return;
			}
			destList = GetSPList(web, destListName);

			XmlNode sourceListNode = null;
			try
			{
				sourceListNode = GetListMetadata(sourceSiteURL,
					sourceListName, listService, out sourceListDescription,
					out sourceListAuthor, out sourceListCreated, out sourceListModified,
					out sourceListType);
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine("Problem accessing the source SharePoint list web service at " + sourceSiteURL + listServiceURL);
				Console.WriteLine(e.Message);
				return;
			}

			if (destList == null)
			{
				newList = true;
				destListDescription = sourceListDescription;
				destListType = sourceListType;
				destList = CreateList(web, destListName, destListDescription, destListType);
			}
			else
			{
				// Check the type of the source and dest, if they don't match, stop.
				if (!sourceListType.Equals(destList.BaseTemplate))
				{
					wrongType = true;
					Console.WriteLine("Destination list " + destListName + " already exists, but of a different type.");
					return;
				}
				else
				{
					Console.WriteLine("Destination list " + destListName + " already exists.");
				}
			}
			// We need to populate listFields even if the list already exists.
			AddFieldsFromXML(destList, sourceListNode.OuterXml.ToString(), listFields, newListFields, reverseListFields);
			try
			{
				XmlNode destListNode = GetListMetadata(destSiteURL,
					destListName, listService, out destListDescription,
					out destListAuthor, out destListCreated, out destListModified,
					out destListType);
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine("Problem accessing the destination SharePoint list web service at " + destSiteURL + listServiceURL);
				Console.WriteLine(e.Message);
				return;
			}
			listService.Url = sourceSiteURL + listServiceURL;
			listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			replacements = aReplacements;

			topLevel = new MBSPListItemMap(this, destFolderPath, true);
		}


		void ParseURLs(string raw_source, string raw_dest_site, string raw_dest_path)
		{
			string dest_path = "";
			string this_folder = "";
			if (raw_dest_path.Length > 0)
			{
				dest_path = raw_dest_path;
				this_folder = dest_path.Substring(dest_path.LastIndexOf('/') + 1);
			}

			sourceSiteURL = GetSiteURL(raw_source);
			sourceListNameURL = MBSPSiteMap.GetListNameURL(raw_source);
			sourceListName = sourceListNameURL;
			destSiteURL = raw_dest_site.Substring(0, raw_dest_site.LastIndexOf('/'));
			destListName = raw_dest_site.Substring(raw_dest_site.LastIndexOf('/') + 1);
			destFolderPath = dest_path;
		}


		void GetRealListName()
		{
			listService.Url = sourceSiteURL + listServiceURL;
			listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			XmlNode lists = listService.GetListCollection();
			foreach (XmlNode child in lists.ChildNodes)
			{
				string name = MBSPSiteMap.GetListNameURL(child.Attributes["DefaultViewUrl"].Value.ToString());
				if (name.Equals(sourceListNameURL))
				{
					sourceListName = child.Attributes["Title"].Value.ToString();
				}
			}
		}


		string GetSiteURL(string url)
		{
			url = Uri.UnescapeDataString(url);
			string siteURL = "";
			if (url.LastIndexOf('/') > 0)
			{
				if (url.Contains(".aspx"))
				{
					string withoutpage = url.Substring(0, url.LastIndexOf('/'));
					siteURL = withoutpage.Substring(0, withoutpage.LastIndexOf('/'));
					if (withoutpage.EndsWith("/Forms"))
					{
						string withoutforms = withoutpage.Substring(0, withoutpage.LastIndexOf('/'));
						// now strip off the list (doclib) name.
						siteURL = withoutforms.Substring(0, withoutforms.LastIndexOf('/'));
					}
				}
				else
				{
					// this should be the form of just the url with the list part in.
					if (url.EndsWith("/"))
					{
						url = url.Substring(0, url.LastIndexOf('/'));
					}
					siteURL = url.Substring(0, url.LastIndexOf('/'));
				}
				if (siteURL.Substring(siteURL.LastIndexOf('/') + 1).Equals("Lists"))
				{
					siteURL = siteURL.Substring(0, siteURL.LastIndexOf('/'));
				}
			}
			return siteURL;
		}


		public SPList CreateList(SPWeb web, string name,
			string description, SPListTemplateType type)
		{
			Guid newList = web.Lists.Add(name, description, type);
			SPList destList = web.Lists[newList];
			// We would really like to set the list metadata here,
			// but there isn't a way to do that.
			return destList;
		}


		static SPWeb GetSPWeb(string site)
		{
			SPSite sc = null;
			SPWeb web = null;
			try
			{
				sc = new SPSite(site);
				web = sc.OpenWeb();
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine(e.Message);
				if (web != null)
				{
					web.Dispose();
				}
				if (sc != null)
				{
					sc.Dispose();
				}
				return null;
			}
			if (sc != null)
			{
				sc.Dispose();
			}
			return web;
		}

		
		SPList GetSPList(SPWeb web, string list)
		{
			SPListCollection allLists = web.Lists;
			foreach (SPList existing_list in allLists)
			{
				if (existing_list.Title.ToLower().Equals(list.ToLower()))
				{
					return existing_list;
				}
			}
			return null;
		}


		XmlNode GetListMetadata(
			string site, 
			string listName, 
			SharePointListsWebService.Lists listService,
			out string listDescription,
			out string listAuthor,
			out DateTime listCreated,
			out DateTime listModified,
			out SPListTemplateType listType)
		{
			listService.Url = site + listServiceURL;
			listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			XmlNode listNode;
			listNode = listService.GetList(listName);
			listDescription = "Migrated List";
			string tempDesc = listNode.Attributes["Description"].Value;
			if (tempDesc.Trim().Length > 0)
			{
				// I'd rather have my default than theirs.
				if (!tempDesc.Trim().Equals("Share a document with the team by adding it to this document library."))
				{
					listDescription = listNode.Attributes["Description"].Value;
				}
			}
			listCreated = DateTime.ParseExact(listNode.Attributes["Created"].Value, "yyyyMMdd HH:mm:ss", CultureInfo.InvariantCulture);
			listModified = DateTime.ParseExact(listNode.Attributes["Modified"].Value, "yyyyMMdd HH:mm:ss", CultureInfo.InvariantCulture);
			listAuthor = MBSPSiteMap.GetLoginNameFromSharePointID(listNode.Attributes["Author"].Value, site);
			listType = GetTypeFromTypeCode(listNode.Attributes["ServerTemplate"].Value);
			return listNode;
		}


		void SetListKeyFieldName(XmlNode fields)
		{
			string firstFieldName = "";
			string firstFieldNameInternal = "";
			string tempInternalName = "";
			string tempDisplayName = "";
			bool titleFieldExists = false;
			bool urlFieldExists = false;
			foreach (XmlNode child in fields)
			{
				bool isID = child.Attributes["Name"].Value.ToString().Equals("ID");
				if (child.Attributes["Name"].Value.ToString().Equals("Title"))
				{
					titleFieldExists = true;
					tempInternalName = child.Attributes["Name"].Value.ToString();
					tempDisplayName = child.Attributes["DisplayName"].Value.ToString();
				}
				if (child.Attributes["Name"].Value.ToString().Equals("URL"))
				{
					urlFieldExists = true;
					tempInternalName = child.Attributes["Name"].Value.ToString();
					tempDisplayName = child.Attributes["DisplayName"].Value.ToString();
				}
				if (firstFieldName.Equals("") && !isID)
				{
					firstFieldName = child.Attributes["DisplayName"].Value.ToString();
					firstFieldNameInternal = child.Attributes["Name"].Value.ToString();
				}
				if (listKeyField.Equals("") && !isID)
				{
					// we want the first required field
					// or if none, a field called Title or URL
					// or if still none, just the first field.
					foreach (XmlAttribute attr in child.Attributes)
					{
						if (attr.Name.Equals("Required"))
						{
							if (child.Attributes["Required"].Value.ToString().Equals("TRUE"))
							{
								listKeyField = child.Attributes["DisplayName"].Value.ToString();
								listKeyFieldInternal = child.Attributes["Name"].Value.ToString();
								break;
							}
						}
					}
				}
			}
			if (listKeyField.Equals(""))
			{
				if (titleFieldExists || urlFieldExists)
				{
					listKeyField = tempDisplayName;
					listKeyFieldInternal = tempInternalName;
				}
				else
				{
					listKeyField = firstFieldName;
					listKeyFieldInternal = firstFieldNameInternal;
				}
			}
		}


		bool ListFieldDisplayNameFound(SPList list, string name)
		{
			foreach (SPField field in list.Fields)
			{
				if (field.Title.Equals(name))
				{
					return true;
				}
			}
			return false;
		}


		bool ListFieldInternalNameFound(SPList list, string name)
		{
			foreach (SPField field in list.Fields)
			{
				if (field.InternalName.Equals(name))
				{
					return true;
				}
			}
			return false;
		}


		/* We have xml from the source list with both the internal name and display name
		 * for each field. However, the Add() method seems to use only the display name
		 * and some of the possible internal names cannot be added. So we need to use the 
		 * display name only, and just map which source internal name gave us which 
		 * display name, and then which display name maps to which new internal name.
		 */
		/*
		Location: Primary Contact
		Author0: Author
		DocIcon: Type
		Review_x0020_by_x0020_Office_x00: Review by Office Staff
		Title: Proposal Name

		Can only add by Display name.
		Can only get data by source internal name.

		Easy, map internal name to display name. Then keep track of new internal name
		for when adding fields.

		But, how to tell if a field exists on the destination?

		Check by display name, if the display name exists, it exists.
		Except - if the internal name is the key field, update the display name but
		only if we're creating the list.
		And - if the display name is Type, the field we were looking for 
		(a user created Type field) does not exist, but only if we're creating the list.
		*/
		void AddFieldFromXML(SPList list, XmlNode node, Hashtable listFields,
			Hashtable newListFields, Hashtable reverseListFields)
		{
			string newInternalName = "";
			string internalName = node.Attributes["Name"].Value.ToString();
			string displayName = node.Attributes["DisplayName"].Value.ToString();
			// store a hashtable of internal name to display name for later.
			listFields.Add(internalName, displayName);
			if (!reverseListFields.ContainsKey(displayName))
			{
				reverseListFields.Add(displayName, internalName);
			}

			// Some fields should just not be copied.
			if (displayName.Length < 1)
			{
				return;
			}
			if (destListType.Equals(SPListTemplateType.PictureLibrary)
				|| destListType.Equals(SPListTemplateType.DocumentLibrary))
			{
				if (internalName.Equals("EncodedAbsThumbnailUrl")
					|| internalName.Equals("EncodedAbsWebImgUrl")
					|| internalName.Equals("SelectedFlag")
					|| displayName.Equals("Name")
					)
				{
					return;
				}
			}
			if (destListType.Equals(SPListTemplateType.DiscussionBoard))
			{
				if (internalName.Equals("Ordering"))
				{
					return;
				}
			}

			bool exists = ListFieldDisplayNameFound(list, displayName);
			if (newList)
			{
				if (internalName.Equals(listKeyFieldInternal)
					&& ListFieldInternalNameFound(list, internalName))
				{
					Console.WriteLine("Updating display name of key field " + internalName + " to " + displayName);
					SPField f = list.Fields.GetFieldByInternalName(internalName);
					f.Title = displayName;
					f.Update(true);
					exists = true;
					if (!reverseNewListFields.ContainsKey(displayName))
					{
						reverseNewListFields.Add(displayName, f.InternalName);
					}
				}
				if (
					((destListType.Equals(SPListTemplateType.DocumentLibrary)
					|| destListType.Equals(SPListTemplateType.PictureLibrary))
					&& ((internalName.Equals("Created") && displayName.Equals("Created Date"))
					 || (internalName.Equals("Modified") && displayName.Equals("Last Modified"))
					 || (internalName.Equals("FileDirRef") && displayName.Equals("URL Dir Name"))
					 || (internalName.Equals("FSObjType") && displayName.Equals("File System Object Type"))))
					|| (destListType.Equals(SPListTemplateType.DiscussionBoard)
					&& (
					(internalName.Equals("Body") && displayName.Equals("Text"))
					|| (internalName.Equals("Author"))
					|| (internalName.Equals("Created"))
					)
					))
				{
					Console.WriteLine("Updating display name of " + internalName + " to " + displayName);
					SPField f = list.Fields.GetFieldByInternalName(internalName);
					f.Title = displayName;
					f.Update(true);
					exists = true;
					if (!reverseNewListFields.ContainsKey(displayName))
					{
						reverseNewListFields.Add(displayName, f.InternalName);
					}
				}
				if (displayName.Equals("Type") 
					&& !destListType.Equals(SPListTemplateType.DocumentLibrary))
				{
					exists = false;
				}
			}

			if (exists)
			{
				bool special = (displayName.Equals("ID") && internalName.Equals("ID"))
					|| (displayName.Equals("owshiddenversion") && internalName.Equals("owshiddenversion"))
					|| (displayName.Equals("Attachments") && internalName.Equals("Attachments"))
					|| (displayName.Equals("Approval Status") && internalName.Equals("_ModerationStatus"))
					|| (displayName.Equals("Approver Comments") && internalName.Equals("_ModerationComments"))
					|| (displayName.Equals("Edit") && internalName.Equals("Edit"))
					|| (displayName.Equals("Select") && internalName.Equals("SelectTitle"))
					|| (displayName.Equals("Order") && internalName.Equals("Order"))
					|| (displayName.Equals("GUID") && internalName.Equals("GUID"))
					|| (displayName.Equals("InstanceID") && internalName.Equals("InstanceID"))
					|| (displayName.Equals("Type") && internalName.Equals("DocIcon"))
					|| (displayName.Equals("View Response") && internalName.StartsWith("DisplayResponse"))
					|| ((destListType.Equals(SPListTemplateType.DocumentLibrary) 
					|| destListType.Equals(SPListTemplateType.PictureLibrary))
					&& ((displayName.Equals("Modified")) && (internalName.Equals("Last_x0020_Modified")))
					|| (displayName.Equals("Created") && internalName.Equals("Created_x0020_Date"))
					|| (displayName.Equals("File System Object Type") && internalName.Equals("FSObjType"))
					|| (displayName.Equals("URL Path") && internalName.Equals("FileRef"))
					|| (displayName.Equals("URL Dir Name") && internalName.Equals("FileDirRef"))
					|| (displayName.Equals("File Size") && internalName.Equals("File_x0020_Size"))
					|| (displayName.Equals("Name") && internalName.Equals("FileLeafRef"))
					|| (displayName.Equals("Virus Status") && internalName.Equals("VirusStatus"))
					|| (displayName.Equals("Shared File Index") && internalName.Equals("_SharedFileIndex"))
					|| (displayName.Equals("Select") && internalName.Equals("SelectFilename"))
					|| (displayName.Equals("Server Relative URL") && internalName.Equals("ServerUrl"))
					|| (displayName.Equals("Encoded Absolute URL") && internalName.Equals("EncodedAbsUrl"))
					|| (displayName.Equals("File Size") && internalName.Equals("FileSizeDisplay"))
					|| (displayName.Equals("Merge") && internalName.Equals("Combine"))
					|| (displayName.Equals("Relink") && internalName.Equals("RepairDocument"))
					|| (internalName.Equals("CheckedOutUserId"))
					|| (internalName.Equals("CheckedOutTitle"))
					|| (internalName.Equals("LinkCheckedOutTitle"))
					|| (internalName.Equals("ImageSize"))
					|| (internalName.Equals("ImageWidth"))
					|| (internalName.Equals("ImageHeight"))
					|| (internalName.Equals("Thumbnail"))
					|| (internalName.Equals("Preview"))
					);
				if (!special && !newListFields.ContainsKey(displayName))
				{
					SPField f = list.Fields.GetField(displayName);
					newListFields.Add(displayName, f.InternalName);
					if (!reverseNewListFields.ContainsKey(displayName))
					{
						reverseNewListFields.Add(displayName, f.InternalName);
					}
				}
			}
			if (!exists)
			{
				Console.Out.WriteLine("Adding field " + displayName + " to list " + list.Title);
				// Modify the XML to have only the display name.
				node.Attributes["Name"].Value = displayName;
				newInternalName = list.Fields.AddFieldAsXml(node.OuterXml);
				newListFields.Add(displayName, newInternalName);
				if (!reverseNewListFields.ContainsKey(displayName))
				{
					reverseNewListFields.Add(displayName, newInternalName);
				}
			}
		}


		public void AddFieldsFromXML(SPList list, string xml, Hashtable listFields,
			Hashtable newListFields, Hashtable reverseListFields)
		{
			XmlReader xmlR = XmlReader.Create(new StringReader(xml));
			XmlDocument doc = new XmlDocument();
			XmlNode node = doc.ReadNode(xmlR);
			XmlNode fields = node.FirstChild;
			SetListKeyFieldName(fields);
			foreach (XmlNode child in fields)
			{
				this.AddFieldFromXML(list, child, listFields, newListFields, reverseListFields);
			}
		}


		public XmlNode GetListItems(string folderName, string listName, string listNameURL, string sourceFolderPath)
		{
			XmlDocument xmlDoc = new System.Xml.XmlDocument();
			XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");
			XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
			XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");

			// if this is the list top level then no folder needs to be specified.
			// TODO: will this work for MOSS2007 lists with folders?
			if (folderName.Length > 0)
			{
				string[] paths = new string[] { listNameURL, sourceFolderPath, folderName };
				string folder = MBSPListMap.CombinePaths(paths);
				ndQueryOptions.InnerXml = "<Folder>" + folder + "</Folder>";
			}
			XmlNode ndListItems;
			listService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			ndListItems = listService.GetListItems(listName, null, ndQuery, ndViewFields, null, ndQueryOptions, null);
			return ndListItems;
		}


		public bool Copy()
		{
			if ((!wrongType) && (topLevel != null))
			{
				topLevel.GetAllSubItems(listService, sourceListName, sourceListNameURL);
				topLevel.CopyData(sourceSiteURL, sourceListNameURL);
				MBSPListViewMap v = new MBSPListViewMap(this);
				return true;
			}
			return false;
		}


		public static string CombinePaths(string [] paths)
		{
			string res = "";
			foreach (string path in paths)
			{
				if ((res.Length > 0) && (path.Length > 0))
				{
					res += "/";
				}
				res += path;
			}
			if ((res.Length > 0) && (res[0].Equals('/')))
			{
				res = res.Substring(1, res.Length - 1);
			}

			return res;
		}


		public SPFolder EnsureFolderPathExists(SPFolder folder, string path)
		{
			if (path.Length < 1)
			{
				return folder;
			}
			DateTime created = sourceListCreated;
			DateTime modified = sourceListModified;
			string author = sourceListAuthor;
			string this_folder = path;
			string below = "";
			int i = path.IndexOf("/");
			if (i == 0)
			{
				this_folder = path.Substring(1);
				i = path.IndexOf("/", 1);
				if (i > 0)
				{
					below = path.Substring(i + 1);
				}
			}
			else if (i > 0)
			{
				this_folder = path.Substring(0, i);
				below = path.Substring(i + 1);
			}
			SPFolder this_level;
			try
			{
				this_level = folder.SubFolders[this_folder];
				//System.Console.Out.WriteLine("");
				//System.Console.Out.WriteLine("Folder " + this_folder + " already exists.");
			}
			catch
			{
				System.Console.Out.WriteLine("");
				System.Console.Out.WriteLine("Creating folder " + this_folder);
				this_level = folder.SubFolders.Add(this_folder);
				this_level.Item["Created"] = created;
				this_level.Item["Modified"] = modified;
				this_level.Item["Author"] = MBSPSiteMap.EnsureAUserExists(author, "", folder.ParentWeb);
				this_level.Item["Modified By"] = MBSPSiteMap.EnsureAUserExists(author, "", folder.ParentWeb);
				this_level.Item.Update();
			}
			if (below.Length > 0)
			{
				return EnsureFolderPathExists(this_level, below);
			}
			return this_level;
		}


		public void Close()
		{
			if (web != null)
			{
				web.Dispose();
			}
		}


		public string ReplaceValues(string value)
		{
			string newValue = value;
			if (replacements.Count > 0)
			{
				ICollection keys = replacements.Keys;
				foreach (object thisKey in keys)
				{
					while (newValue.Contains(thisKey.ToString()))
					{
						Console.WriteLine("");
						Console.WriteLine("Replacing " + thisKey.ToString() + " in " + newValue);
						newValue = newValue.Replace(thisKey.ToString(), replacements[thisKey].ToString());
					}
				}
			}
			return newValue;
		}


		// From http://www.sharepointblogs.com/marwantarek/archive/2007/08/12/list-definitions-type-and-basetype.aspx
		public static SPListTemplateType GetTypeFromTypeCode(string typeCode)
		{
			switch (typeCode)
			{
				case "100":
					return SPListTemplateType.GenericList;
				case "101":
					return SPListTemplateType.DocumentLibrary;
				case "102":
					return SPListTemplateType.Survey;
				case "103":
					return SPListTemplateType.Links;
				case "104":
					return SPListTemplateType.Announcements;
				case "105":
					return SPListTemplateType.Contacts;
				case "106":
					return SPListTemplateType.Events;
				case "107":
					return SPListTemplateType.Tasks;
				case "108":
					return SPListTemplateType.DiscussionBoard;
				case "109":
					return SPListTemplateType.PictureLibrary;
				case "110":
					return SPListTemplateType.DataSources;
				case "120":
					return SPListTemplateType.CustomGrid;
			}
			return SPListTemplateType.DocumentLibrary;
		}


		public string GetSourceSiteURL()
		{
			return sourceSiteURL;
		}


		public string GetSourceListName()
		{
			return sourceListName;
		}


		public SPListTemplateType GetSourceListType()
		{
			return sourceListType;
		}
	}
}
