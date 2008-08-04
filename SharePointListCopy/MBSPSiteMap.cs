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
using System.DirectoryServices;
using System.IO;
using System.Xml;
using Microsoft.SharePoint;


namespace SharePointListCopy
{
	class MBSPSiteMap
	{
		static Hashtable userIDCache = new Hashtable();
		static Hashtable userCache1 = new Hashtable();
		static Hashtable userCache2 = new Hashtable();
		static string siteURL1 = "";
		static string siteURL2 = "";
		public static string defaultDomain = Environment.UserDomainName;


		public static bool CreateLocalAccount(string loginName, string fullName, out string domain)
		{
			domain = Environment.MachineName;
			try
			{
				DirectoryEntry AD = new DirectoryEntry("WinNT://" +
									Environment.MachineName + ",computer");
				DirectoryEntry NewUser = AD.Children.Add(loginName, "user");
				NewUser.Invoke("SetPassword", new object[] { "#12%tF11@7345Abc" });
				NewUser.Invoke("Put", new object[] { "Description", "SharePoint migrated user" });
				NewUser.Invoke("Put", new object[] { "FullName", fullName });
				int ADS_UF_PASSWD_CANT_CHANGE = 0x000000040;
				int ADS_UF_DONT_EXPIRE_PASSWD = 0x00010000;
				int ADS_UF_ACCOUNTDISABLE = 0x0002;
				int combinedFlag = 0;
				combinedFlag = ADS_UF_DONT_EXPIRE_PASSWD | ADS_UF_PASSWD_CANT_CHANGE;
				combinedFlag = combinedFlag | ADS_UF_ACCOUNTDISABLE;
				NewUser.Invoke("Put", new Object[] { "userFlags", combinedFlag });
				NewUser.CommitChanges();
				return true;
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error creating local account: " + ex.Message);
				return false;
			}
		}


		public static string EnsureAUserExists(string loginName, string fullName, SPWeb web)
		{
			if (userIDCache.Count > 0)
			{
				ICollection userkeys = userIDCache.Keys;
				foreach (object userkey in userkeys)
				{
					if (userkey.ToString().Equals(loginName))
					{
						return userIDCache[loginName].ToString();
					}
				}
			}
			Console.WriteLine("");
			Console.WriteLine("Looking for user " + loginName);
			SPUser User = null;
			try
			{
				User = web.EnsureUser(loginName);
			}
			catch
			{
				//Program.logFile.WriteLine(loginName);
				//Program.logFile.Flush();
				int sep = loginName.IndexOf('\\');
				string shortLoginName = loginName.Substring(sep + 1);
				try
				{
					User = web.EnsureUser(defaultDomain + "\\" + shortLoginName);
				}
				catch
				{
					try
					{
						User = web.EnsureUser(shortLoginName);
					}
					catch
					{
						string domain = "";
						bool r = CreateLocalAccount(shortLoginName, fullName, out domain);
						if (r)
						{
							Console.WriteLine("");
							Console.WriteLine("Creating user " + fullName + " (" + shortLoginName + ")");
							web.SiteUsers.Add(domain + "\\" + shortLoginName, "", fullName, "");
							User = web.EnsureUser(shortLoginName);
						}
					}
				}
			}
			int userid = 1;
			if (User != null)
			{
				userid = User.ID;
			}
			userIDCache.Add(loginName, userid);
			Console.WriteLine("Found user " + loginName + " and cached user ID " + userid);
			return userid.ToString();
		}


		public static string GetFullNameFromSharePointID(string userString)
		{
			int i = userString.IndexOf(";#");
			string name = "";
			if (i > 0)
			{
				name = userString.Substring(userString.IndexOf(";#") + 2);
			}
			return name;
		}


		// Look up the login name from the source SharePoint user ID.
		public static string GetLoginNameFromSharePointID(string userString, string siteURL)
		{
			if (siteURL1.Length < 1)
			{
				siteURL1 = siteURL;
			}
			else
			{
				if (siteURL2.Length < 1)
				{
					if (!siteURL.Equals(siteURL1))
					{
						siteURL2 = siteURL;
					}
				}
			}
			int i = userString.IndexOf(";#");
			string ID = "";
			if (i == -1)
			{
				ID = userString;
			}
			else
			{
				ID = userString.Substring(0, userString.IndexOf(";#"));
			}
			if (siteURL.Equals(siteURL1))
			{
				if (userCache1.Count > 0)
				{
					ICollection userkeys = userCache1.Keys;
					foreach (object userkey in userkeys)
					{
						if (userkey.ToString().Equals(ID))
						{
							return userCache1[ID].ToString();
						}
					}
				}
			}
			else if (siteURL.Equals(siteURL2))
			{
				if (userCache2.Count > 0)
				{
					ICollection userkeys = userCache2.Keys;
					foreach (object userkey in userkeys)
					{
						if (userkey.ToString().Equals(ID))
						{
							return userCache2[ID].ToString();
						}
					}
				}
			}
			SharePointUserGroupWebService.UserGroup userService = new SharePointUserGroupWebService.UserGroup();
			userService.Url = siteURL + "/_vti_bin/UserGroup.asmx";
			userService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			XmlNode users = userService.GetUserCollectionFromSite();
			String userQuery = "//*[@*]";
			XmlNodeList userList = users.SelectNodes(userQuery);
			string loginName = "";
			for (int u = 0; u < userList.Count; u++)
			{
				if (userList[u].Attributes["ID"].Value.Equals(ID))
				{
					loginName = userList[u].Attributes["LoginName"].Value;
					if (siteURL.Equals(siteURL1))
					{
						userCache1.Add(ID, loginName);
					}
					else if (siteURL.Equals(siteURL2))
					{
						userCache2.Add(ID, loginName);
					}
					return loginName;
				}
			}
			return "";
		}


		public static byte[] ByteArrayFromFilePath(string file)
		{
			FileInfo info = new FileInfo(file);
			long byteCount = info.Length;
			FileStream stream = new FileStream(file, FileMode.Open);
			BinaryReader br = new BinaryReader(stream);
			byte[] data = br.ReadBytes((int)byteCount);
			br.Close();
			stream.Close();
			return data;
		}


		public static string GetListNameURL(string url)
		{
			url = Uri.UnescapeDataString(url);
			string listURL = "";
			if (url.LastIndexOf('/') > 0)
			{
				if (url.Contains(".aspx"))
				{
					string withoutpage = url.Substring(0, url.LastIndexOf('/'));
					listURL = withoutpage.Substring(withoutpage.LastIndexOf('/') + 1);
					if (withoutpage.EndsWith("/Forms"))
					{
						string withoutforms = withoutpage.Substring(0, withoutpage.LastIndexOf('/'));
						// now get the list (doclib) name.
						listURL = withoutforms.Substring(withoutforms.LastIndexOf('/') + 1);
					}
				}
				else
				{
					// this should be the form of just the url with the list part in.
					if (url.EndsWith("/"))
					{
						url = url.Substring(0, url.LastIndexOf('/'));
					}
					listURL = url.Substring(url.LastIndexOf('/') + 1);
				}
			}
			return listURL;
		}


		public static ArrayList GetSiteLists(string sourceSiteURL, string destSiteURL)
		{
			string source = "";
			string dest_site = "";
			string dest_path = "";
			string listNameURL = "";
			ArrayList r = new ArrayList();
			int startOfFQDN = sourceSiteURL.IndexOf("//") + 2;
			int endOfFQDN = sourceSiteURL.IndexOf("/", startOfFQDN);
			if (endOfFQDN.Equals(-1))
			{
				endOfFQDN = sourceSiteURL.Length;
			}
			string sourceSiteURLBase = sourceSiteURL.Substring(0, endOfFQDN);
			if (!destSiteURL.EndsWith("/"))
			{
				destSiteURL += "/";
			}
			SharePointSiteDataWebService.SiteData siteService = new SharePointSiteDataWebService.SiteData();
			siteService.Url = sourceSiteURL + "/_vti_bin/SiteData.asmx";
			siteService.Credentials = System.Net.CredentialCache.DefaultCredentials;
			SharePointSiteDataWebService._sList[] lists;
			siteService.GetListCollection(out lists);
			foreach (SharePointSiteDataWebService._sList list in lists)
			{
				source = sourceSiteURLBase + list.DefaultViewUrl;
				listNameURL = GetListNameURL(source);
				dest_site = destSiteURL + listNameURL;
				string[] bits = { source, dest_site, dest_path };
				Console.WriteLine("");
				Console.WriteLine("Adding site list: " + bits[0] + " copying to " + bits[1]);
				r.Add(bits);
			}
			return r;
		}
	}
}
