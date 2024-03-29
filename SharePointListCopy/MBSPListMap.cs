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
		SharePointSiteDataWebService.SiteData siteDataService = new SharePointSiteDataWebService.SiteData();
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
		bool sourceListEnableVersions = false;
		bool destListEnableVersions = false;
		string sourceListID = "";
		string destListID = "";
		public Hashtable listFields = new Hashtable();
		public Hashtable reverseListFields = new Hashtable();
		public Hashtable newListFields = new Hashtable();
		public Hashtable reverseNewListFields = new Hashtable();
		bool newList = false;
		string listServiceURL = "/_vti_bin/Lists.asmx";
		string siteDataServiceURL = "/_vti_bin/SiteData.asmx";
		public string listKeyField = "";
		string listKeyFieldInternal = "";
		bool wrongType = false;
		MBSPListItemMap topLevel;
		Hashtable replacements;
		bool hasLookupFields = false;
		bool isUnsupportedListType = false;
		bool dependsOnMissingLists = false;
		XmlNode sourceListNode = null;


		public MBSPListMap(string raw_source, string raw_dest_site, string raw_dest_path,
			Hashtable aReplacements)
		{
			replacements = aReplacements;
			Console.WriteLine("");
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
			listService.Credentials = Program.getSourceCredentials();
			GetRealListName();
			sourceListNode = null;
			try
			{
				sourceListNode = GetListMetadata(sourceSiteURL,
					sourceListName, listService, Program.getSourceCredentials(), out sourceListDescription,
					out sourceListAuthor, out sourceListCreated, out sourceListModified,
					out sourceListType, out sourceListEnableVersions, out sourceListID);
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine("Problem accessing the source SharePoint list web service at " + sourceSiteURL + listServiceURL);
				Console.WriteLine(e.Message);
				return;
			}

			if (sourceListType == SPListTemplateType.Agenda
				|| sourceListType == SPListTemplateType.MeetingUser
				|| sourceListType == SPListTemplateType.MeetingObjective)
			{
				// Trying to add these list types to the SPWeb only returns "Invalid list template."
				isUnsupportedListType = true;
			}

			XmlNamespaceManager nsmgr = new XmlNamespaceManager(sourceListNode.OwnerDocument.NameTable);
			nsmgr.AddNamespace("soap", "http://schemas.microsoft.com/sharepoint/soap/");
			string xpath = "/soap:Fields/soap:Field[@Type='Lookup']";
			XmlNodeList fields = sourceListNode.SelectNodes(xpath, nsmgr);
			foreach (XmlNode field in fields)
			{
				// Document libraries have lots of Lookup fields, with List set to "Docs".
				if (!field.Attributes["List"].Value.Equals("Docs"))
				{
					hasLookupFields = true;
					break; // One is enough.
				}
			}
		}


		public bool Init()
		{
			if (Program.onlyAddNewFilesInDoclibs)
			{
				if (!(sourceListType.Equals(SPListTemplateType.DocumentLibrary)
					|| sourceListType.Equals(SPListTemplateType.PictureLibrary))
					)
				{
					Console.WriteLine("Skipping this list because --only-add-new-files-in-doclibs only allows document and picture libraries.");
					return false;
				}
			}
			web = GetSPWeb(destSiteURL);
			if (web == null)
			{
				return false;
			}
			if (destListName.Length < 1)
			{
				destListName = sourceListName;
			}
			destList = GetSPList(web, destListName);


			if (destList == null)
			{
				newList = true;
				destListDescription = sourceListDescription;
				destListType = sourceListType;
				destListEnableVersions = sourceListEnableVersions;
				destList = CreateList(web, destListName, destListDescription, destListType, destListEnableVersions);
			}
			else
			{
				destListType = destList.BaseTemplate;
				// Check the type of the source and dest, if they don't match, stop.
				if (!sourceListType.Equals(destList.BaseTemplate))
				{
					wrongType = true;
					Console.WriteLine("Destination list " + destListName + " already exists, but of a different type.");
					return false;
				}
				else
				{
					Console.WriteLine("Destination list " + destListName + " already exists.");
				}
			}
			// We need to populate listFields even if the list already exists.
			bool r = AddFieldsFromXML(destList, sourceListNode.OuterXml.ToString(), listFields, newListFields, reverseListFields);
			if (r.Equals(false))
			{
				return r;
			}
			try
			{
				XmlNode destListNode = GetListMetadata(destSiteURL,
					destListName, listService, System.Net.CredentialCache.DefaultCredentials, out destListDescription,
					out destListAuthor, out destListCreated, out destListModified,
					out destListType, out destListEnableVersions, out destListID);
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine("Problem accessing the destination SharePoint list web service at " + destSiteURL + listServiceURL);
				Console.WriteLine(e.Message);
				return false;
			}
			listService.Url = sourceSiteURL + listServiceURL;
			listService.Credentials = Program.getSourceCredentials();

			topLevel = new MBSPListItemMap(this, destFolderPath, true);
			return true;
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
			listService.Credentials = Program.getSourceCredentials();
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
			string description, SPListTemplateType type, bool enableVersions)
		{
			Guid newList = web.Lists.Add(name, description, type);
			SPList destList = web.Lists[newList];
			if (destList.BaseTemplate != SPListTemplateType.Survey)
			{
				if (Program.forceVersioning)
				{
					enableVersions = true;
				}
				destList.EnableVersioning = enableVersions;
				destList.Update();
			}
			destList.OnQuickLaunch = true;
			destList.Update();
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
				if (!Program.singleList && !web.Url.Equals(site))
				{
				    if (Program.newSiteTemplate != "")
				    {
						SPWebTemplateCollection Templates = web.GetAvailableWebTemplates(1033);
						SPWebTemplate siteTemplate = Templates[Program.newSiteTemplate];
						string site_name = site.Substring(site.LastIndexOf('/') + 1);
						if (Program.beVerbose)
						{
							Console.WriteLine("");
							Console.WriteLine("");
							Console.WriteLine("Creating site " + site_name + " from template " + Program.newSiteTemplate);
							Console.WriteLine("");
							Console.WriteLine("");
						}
						web.Webs.Add(site_name, site_name, "", 1033, siteTemplate, false, false);
						sc = new SPSite(site);
						web = sc.OpenWeb();
				    }
					else if (Program.createBlankSite)
					{
						SPWebTemplateCollection Templates = web.GetAvailableWebTemplates(1033);
						SPWebTemplate siteTemplate = Templates["STS#1"];
						string site_name = site.Substring(site.LastIndexOf('/') + 1);
						if (Program.beVerbose)
						{
							Console.WriteLine("");
							Console.WriteLine("");
							Console.WriteLine("Creating blank site " + site_name);
							Console.WriteLine("");
							Console.WriteLine("");
						}
						web.Webs.Add(site_name, site_name, "", 1033, siteTemplate, false, false);
						sc = new SPSite(site);
						web = sc.OpenWeb();
					}
					else
					{
						Exception e = new Exception("The destination site " + site + " does not exist. Please create it first.");
						throw e;
					}
				}
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
			System.Net.ICredentials credentials,
			out string listDescription,
			out string listAuthor,
			out DateTime listCreated,
			out DateTime listModified,
			out SPListTemplateType listType,
			out bool listEnableVersions,
			out string listID)
		{
			listService.Url = site + listServiceURL;
			listService.Credentials = credentials;
			XmlNode listNode;
			listNode = listService.GetList(listName);
			listDescription = "Migrated List";
			bool enableVersions = false;
			XmlNode versions = listNode.Attributes.GetNamedItem("EnableVersioning");
			if (versions != null)
			{
				enableVersions = System.Convert.ToBoolean(versions.Value.ToString());
			}
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
			listAuthor = MBSPSiteMap.GetLoginNameFromSharePointID(listNode.Attributes["Author"].Value, site, credentials);
			listType = GetTypeFromTypeCode(listNode.Attributes["ServerTemplate"].Value);
			listEnableVersions = enableVersions;
			listID = listNode.Attributes["ID"].Value;
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
		bool AddFieldFromXML(SPList list, XmlNode node, Hashtable listFields,
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
				return true;
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
					return true;
				}
			}
			if (destListType.Equals(SPListTemplateType.DiscussionBoard))
			{
				if (internalName.Equals("Ordering"))
				{
					return true;
				}
			}
			if (destListType.Equals(SPListTemplateType.IssueTracking))
			{
				if (internalName.Equals("ID")
					|| internalName.Equals("IssueID")
					|| internalName.Equals("RemoveRelatedID")
					|| internalName.Equals("LinkIssueIDNoMenu"))
				{
					return true;
				}
			}

			bool exists = ListFieldDisplayNameFound(list, displayName);
			bool exists2 = ListFieldInternalNameFound(list, internalName);
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
				)
				|| (destListType.Equals(SPListTemplateType.IssueTracking)
				&& (internalName.Equals("InstanceID"))
				)
				|| (destListType.Equals(SPListTemplateType.Events)
				&& ((internalName.Equals("EventDate"))
				|| (internalName.Equals("EndDate"))
				))
				)
			{
				if (Program.beVerbose)
				{
					Console.WriteLine("Updating display name of " + internalName + " to " + displayName);
				}
				SPField f = list.Fields.GetFieldByInternalName(internalName);
				f.Title = displayName;
				f.Update(true);
				exists = true;
				if (!reverseNewListFields.ContainsKey(displayName))
				{
					reverseNewListFields.Add(displayName, f.InternalName);
				}
			}
			if (newList)
			{
				if (internalName.Equals(listKeyFieldInternal)
					&& ListFieldInternalNameFound(list, internalName))
				{
					if (Program.beVerbose)
					{
						Console.WriteLine("Updating display name of key field " + internalName + " to " + displayName);
					}
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
				bool special = (displayName.Equals("owshiddenversion") && internalName.Equals("owshiddenversion"))
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
				// If this list exists, skip the ID because it's likely to conflict.
				if (!newList)
				{
					special = special || (displayName.Equals("ID") && internalName.Equals("ID"));
				}
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
				if (Program.beVerbose)
				{
					Console.Out.WriteLine("Adding field " + displayName + " to list " + list.Title);
				}
				// Modify the XML to have only the display name.
				node.Attributes["Name"].Value = displayName;
				// Switch list IDs for the Lookup field
				if (node.Attributes["Type"].Value == "Lookup")
				{
					if (Program.listIDs.ContainsKey(node.Attributes["List"].Value))
					{
						node.Attributes["List"].Value = Program.listIDs[node.Attributes["List"].Value].ToString();
					}
					else
					{
						string listName = MBSPSiteMap.GetListNameFromID(sourceSiteURL, node.Attributes["List"].Value);
						try
						{
							string id = list.ParentWeb.Lists[listName].ID.ToString();
							node.Attributes["List"].Value = id;
						}
						catch
						{
							// If the list with the field this is trying to reference doesn't yet exist,
							// leave this list till even later.
							Console.WriteLine("");
							Console.WriteLine("*** WARNING: List " + list.Title + " depends on another list (" + listName + ") that is missing or not yet copied, so this half done list will be removed. This list might not be copied in this run.");
							list.Delete();
							dependsOnMissingLists = true;
							return false;
						}
					}
				}
				newInternalName = list.Fields.AddFieldAsXml(node.OuterXml);
				newListFields.Add(displayName, newInternalName);
				if (!reverseNewListFields.ContainsKey(displayName))
				{
					reverseNewListFields.Add(displayName, newInternalName);
				}
			}
			return true;
		}


		public bool AddFieldsFromXML(SPList list, string xml, Hashtable listFields,
			Hashtable newListFields, Hashtable reverseListFields)
		{
			XmlReader xmlR = XmlReader.Create(new StringReader(xml));
			XmlDocument doc = new XmlDocument();
			XmlNode node = doc.ReadNode(xmlR);
			XmlNode fields = node.FirstChild;
			SetListKeyFieldName(fields);
			bool r = true;
			foreach (XmlNode child in fields)
			{
				r = this.AddFieldFromXML(list, child, listFields, newListFields, reverseListFields);
				if (r.Equals(false))
				{
					return r;
				}
			}
			return true;
		}


		public XmlNode GetListItems(string folderName, string listName, string listNameURL, string sourceFolderPath, System.Net.ICredentials credentials)
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
			listService.Credentials = credentials;
			ndListItems = listService.GetListItems(listName, null, ndQuery, ndViewFields, System.Int32.MaxValue.ToString(), ndQueryOptions, null);
			return ndListItems;
		}


		public XmlNode GetAllListItems()
		{
			if (sourceListID.Length < 1)
			{
				listService.Credentials = Program.getSourceCredentials();
				XmlNode allLists = listService.GetListCollection();
				String xpq = "//*[@*]"; //get all nodes
				XmlNodeList allNodes = allLists.SelectNodes(xpq);
				for (int i = 0; i < allNodes.Count; i++)
				{
					XmlNode listItemNode = allNodes[i];
					if (sourceListName.Equals(listItemNode.Attributes["Title"].Value.ToString()))
					{
						sourceListID = listItemNode.Attributes["ID"].Value.ToString();
					}
				}
			}

			siteDataService.Url = sourceSiteURL + siteDataServiceURL;
			siteDataService.Credentials = Program.getSourceCredentials();
			string xml = siteDataService.GetListItems(sourceListID, "", "", System.UInt32.MaxValue);
			XmlReader xmlReader = XmlReader.Create(new StringReader(xml));
			XmlDocument xmlDoc = new XmlDocument();
			XmlNode node = xmlDoc.ReadNode(xmlReader);
			//Console.WriteLine(xml);
			return node;
		}


		public bool Copy()
		{
			if ((!wrongType) && (topLevel != null))
			{
				if (Program.preferFolderMetadata)
				{
					topLevel.GetAllSubItems(listService, sourceListName, sourceListNameURL);
				}
				else
				{
					topLevel.GetMoreSubItems(listService, sourceListName, sourceListNameURL);
				}
				topLevel.CopyData(sourceSiteURL, sourceListNameURL);
				if (Program.onlyAddNewFilesInDoclibs)
				{
					Console.WriteLine("");
					Console.WriteLine("Skipping copy of Views.");
				}
				else
				{
					MBSPListViewMap v = new MBSPListViewMap(this);
				}
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
				if (Program.beVerbose)
				{
					System.Console.Out.WriteLine("");
					System.Console.Out.WriteLine("Creating folder " + this_folder);
				}
				try
				{
					// Failed attempt to improve folder creation.
					//this_level = folder.SubFolders.Add("splistcptempnewfolder");
					//this_level.Item["Name"] = this_folder;

					// I've posted about this crap at:
					// http://stackoverflow.com/questions/1040804
					//string fullFolderURL = folder.ParentWeb.Url + folder.ServerRelativeUrl;
					//SPList list = folder.ParentWeb.Lists.GetList(folder.ParentListId, true);
					//SPListItem newFolder = list.Items.Add(fullFolderURL, SPFileSystemObjectType.Folder, this_folder);
					//newFolder.Update();
					//this_level = newFolder.Folder;

					// Original code to create folders, not recommended for WSS 3.0 apparently:
					// http://vspug.com/stevekay72/2007/08/16/create-sub-folders-in-lists-programmatically/
					// Leaving this in for now, since early failure is better than creating folders
					// I don't want as the newer code above will do.
					this_level = folder.SubFolders.Add(this_folder);

					this_level.Item["Created"] = created;
					this_level.Item["Modified"] = modified;
					this_level.Item["Author"] = MBSPSiteMap.EnsureAUserExists(author, "", folder.ParentWeb);
					this_level.Item["Modified By"] = MBSPSiteMap.EnsureAUserExists(author, "", folder.ParentWeb);
					this_level.Item.Update();
				}
				catch (Exception e)
				{
					Console.WriteLine("Unable to create folder " + this_folder);
					throw e;
				}
			}
			if (below.Length > 0)
			{
				return EnsureFolderPathExists(this_level, below);
			}
			return this_level;
		}


		public SPFolder FindFolderFromPath(SPFolder folder, string path)
		{
			if (path.Length < 1)
			{
				return folder;
			}
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
				return null;
			}
			if (below.Length > 0)
			{
				return FindFolderFromPath(this_level, below);
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
						if (Program.beVerbose)
						{
							Console.WriteLine("");
							Console.WriteLine("Replacing " + thisKey.ToString() + " in " + newValue);
						}
						newValue = newValue.Replace(thisKey.ToString(), replacements[thisKey].ToString());
					}
				}
			}
			return newValue;
		}


		// From http://www.sharepointblogs.com/marwantarek/archive/2007/08/12/list-definitions-type-and-basetype.aspx
		// And http://www.sharepointblogs.com/bobbyhabib/archive/2007/09/26/list-types-amp-list-internal-values-available-in-moss-2007.aspx
		// And http://www.codeproject.com/KB/dotnet/QueriesToAnalyzeSPUsage.aspx
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
				case "111":
					return SPListTemplateType.WebTemplateCatalog;
				case "113":
					return SPListTemplateType.WebPartCatalog;
				case "114":
					return SPListTemplateType.ListTemplateCatalog;
				case "115":
					return SPListTemplateType.XMLForm;
				case "120":
					return SPListTemplateType.CustomGrid;
				case "200":
					return SPListTemplateType.Meetings;
				case "201":
					return SPListTemplateType.Agenda;
				case "202":
					return SPListTemplateType.MeetingUser;
				case "204":
					return SPListTemplateType.Decision;
				case "207":
					return SPListTemplateType.MeetingObjective;
				case "210":
					return SPListTemplateType.TextBox;
				case "211":
					return SPListTemplateType.ThingsToBring;
				case "212":
					return SPListTemplateType.HomePageLibrary;
				case "1100":
					return SPListTemplateType.IssueTracking;
			}
			// This is not helpful.
			//return SPListTemplateType.DocumentLibrary;
			Exception e = new Exception("SharePoint list template type code " + typeCode.ToString() + " is not yet supported.");
			throw e;
		}


		public string GetSourceSiteURL()
		{
			return sourceSiteURL;
		}


		public string GetSourceListName()
		{
			return sourceListName;
		}


		public string GetSourceListNameURL()
		{
			return sourceListNameURL;
		}


		public SPListTemplateType GetSourceListType()
		{
			return sourceListType;
		}


		public bool GetDestListEnableVersions()
		{
			return destListEnableVersions;
		}


		public bool HasLookupFields()
		{
			return hasLookupFields;
		}


		public bool IsUnsupportedListType()
		{
			return isUnsupportedListType;
		}


		public bool DependsOnMissingLists()
		{
			return dependsOnMissingLists;
		}


		public string SourceListID()
		{
			return sourceListID;
		}


		public string DestListID()
		{
			return destListID;
		}

		public SPListTemplateType SourceListType()
		{
			return sourceListType;
		}
	}
}
