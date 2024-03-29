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

/*
 * Note: this tool must be run on the actual destination server.
 * Note: may require changes on destination as per http://support.microsoft.com/kb/896861.
 *
 * TODO:
 * - support MOSS2007 as source
 * - support lists with folders
 * - http://blog.krichie.com/2007/01/30/traversing-sharepoint-list-folder-hierarchies/
*/

using System;
using System.IO;
using System.Collections;
using System.Xml;
using System.Globalization;
using Microsoft.SharePoint;

namespace SharePointListCopy
{
	class Program
	{
		// We are only guessing about what is a duplicate, and most often
		// there will not be an existing list, so default is copy everything.
		public static bool avoidDuplicates = false;
		public static string tempFilePath = Path.GetTempPath() + "/splistcp";
		static Hashtable options = new Hashtable();
		static Hashtable optionValues = new Hashtable();
		static ArrayList lists = new ArrayList();
		static ArrayList redoLists = new ArrayList();
		static ArrayList redoLists2 = new ArrayList();
		static Hashtable replacements = new Hashtable();
		public static bool singleList = false;
		public static bool createBlankSite = false;
		public static string newSiteTemplate = "";
		public static string sourceCredentialsDomain = "";
		public static string sourceCredentialsUsername = "";
		public static string sourceCredentialsPassword = "";
		public static bool skipOldVersions = false;
		public static bool beVerbose = false;
		public static Hashtable listIDs = new Hashtable();
		public static bool forceVersioning = false;
		public static bool versionsUseUSDates = false;
		public static bool preferFolderMetadata = false;
		public static bool onlyAddNewFilesInDoclibs = false;
		//public static string logFilePath = "";
		//public static StreamWriter logFile;

		static void Main(string[] args)
		{
			try
			{
				/*
				string now = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString();
				now += DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString();
				now += DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString();
				logFilePath = "C:/splistcp-" + now + ".log";
				logFile = new StreamWriter(logFilePath);
				*/

				if (!Init(args))
				{
					return;
				}
				foreach (string[] list in lists)
				{
					MBSPListMap listMap;
					listMap = new MBSPListMap(list[0], list[1], list[2], replacements);
					if (listMap.HasLookupFields())
					{
						if (Program.beVerbose)
						{
							Console.WriteLine("List " + listMap.GetSourceListName() + " has at least one Lookup field, leaving till last.");
						}
						redoLists.Add(list);
					}
					else if (listMap.IsUnsupportedListType())
					{
						Console.WriteLine("WARNING: List " + listMap.GetSourceListName() + " is a type that cannot yet be copied by this software.");
					}
					else
					{
						if (listMap.Init())
						{
							listIDs[listMap.SourceListID()] = listMap.DestListID();
							listMap.Copy();
						}
						else if (listMap.DependsOnMissingLists())
						{
							redoLists.Add(list);
						}
					}
					listMap.Close();
				}
				foreach (string[] list in redoLists)
				{
					MBSPListMap listMap;
					listMap = new MBSPListMap(list[0], list[1], list[2], replacements);
					if (listMap.Init())
					{
						listIDs[listMap.SourceListID()] = listMap.DestListID();
						listMap.Copy();
					}
					else if (listMap.DependsOnMissingLists())
					{
						redoLists2.Add(list);
					}
					listMap.Close();
				}
				foreach (string[] list in redoLists2)
				{
					MBSPListMap listMap;
					listMap = new MBSPListMap(list[0], list[1], list[2], replacements);
					if (listMap.Init())
					{
						listIDs[listMap.SourceListID()] = listMap.DestListID();
						listMap.Copy();
					}
					listMap.Close();
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine("** ERROR: Something has caused this to fail completely:");
				Console.WriteLine(e.Message);
			}
			finally
			{
				//logFile.Close();
			}
			//System.Threading.Thread.Sleep(5000);
			//Console.ReadLine();
		}


		static bool Init(string[] args)
		{
			Rev r = new Rev();
			string svnrev = r.svnrev;
			System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
			int major = a.GetName().Version.Major;
			int minor = a.GetName().Version.Minor;
			int build = a.GetName().Version.Build;
			int revision = a.GetName().Version.Revision;
			string version = major.ToString() + "." + minor.ToString() + "." + build.ToString();
			Console.WriteLine("SharePointListCopy v" + version + " (r" + svnrev + "). Copyright 2008 by Michael Baltaks.");

			options.Add("--avoid-duplicates", "Use to guess at a key field and avoid copying items with the same key value.");
			options.Add("--replacements-file", "Path to a file of values to look for, and the replacements.");
			options.Add("--replacements-separator", "Separator used between values to find and replace in the replacements file. Default is ||.");
			options.Add("--lists-file", "Path to a file of source and destination list URLs with optional folder path. Replaces the command line URLs.");
			options.Add("--lists-separator", "Separator used between values in the lists file. Default is ||.");
			options.Add("--default-domain", "Look for users in this domain, defaults to domain of current user.");
			options.Add("--single-list", "Will copy only the list at the URL specified.");
			options.Add("--create-blank-site", "Creates a site at the destination from the Blank Site template, if no such site exists. Will not work with --single-list option.");
			options.Add("--create-site-from-template", "Creates a site at the destination from the specified template, if no such site exists. Will not work with --single-list option. Overrides --create-blank-site.");
			options.Add("--source-credentials-domain", "Provide a domain name as part of the credentials used to access the source sharepoint site.");
			options.Add("--source-credentials-username", "Provide a username as part of the credentials used to access the source sharepoint site.");
			options.Add("--source-credentials-password", "Provide a password as part of the credentials used to access the source sharepoint site.");
			options.Add("--skip-old-versions", "Don't bother looking up and copying across old versions, just keep the most recent version.");
			options.Add("--verbose", "Print extra operational messages about progress.");
			options.Add("--force-versioning", "Ensure that versioning is turned on for the destination list(s), no matter what setting the original list(s) used.");
			options.Add("--versions-use-us-dates", "Specify that the date format for versions is in US date format.");
			options.Add("--prefer-folder-metadata", "Use the method of finding document library contents that will find folder metadata, but which depends on what is available in the default view.");
			options.Add("--only-add-new-files-in-doclibs", "Only look at document libraries and picture libraries, and only copy files that don't exist in the destination.");
			//options.Add("--doclibs-only", "");
			//options.Add("--lists-only", "");

			string raw_source, raw_dest_site, raw_dest_path;
			if (!HandleArgs(args, out raw_source, out raw_dest_site, out raw_dest_path))
			{
				return false;
			}
			string listsSeparator = "||";
			if (optionValues.ContainsKey("--lists-separator"))
			{
				listsSeparator = optionValues["--lists-separator"].ToString();
			}
			avoidDuplicates = optionValues.ContainsKey("--avoid-duplicates");
			string replacementsFile = "";
			if (optionValues.ContainsKey("--replacements-file"))
			{
				replacementsFile = optionValues["--replacements-file"].ToString();
			}
			string replacementsSeparator = "||";
			if (optionValues.ContainsKey("--replacements-separator"))
			{
				replacementsSeparator = optionValues["--replacements-separator"].ToString();
			}
			replacements = GetReplacements(replacementsFile, replacementsSeparator);
			if (optionValues.ContainsKey("--default-domain"))
			{
				MBSPSiteMap.defaultDomain = optionValues["--default-domain"].ToString();
			}
			if (optionValues.ContainsKey("--create-blank-site"))
			{
				createBlankSite = true;
			}
			if (optionValues.ContainsKey("--create-site-from-template"))
			{
				newSiteTemplate = optionValues["--create-site-from-template"].ToString();
			}
			if (optionValues.ContainsKey("--source-credentials-domain"))
			{
				sourceCredentialsDomain = optionValues["--source-credentials-domain"].ToString();
			}
			if (optionValues.ContainsKey("--source-credentials-username"))
			{
				sourceCredentialsUsername = optionValues["--source-credentials-username"].ToString();
			}
			if (optionValues.ContainsKey("--source-credentials-password"))
			{
				sourceCredentialsPassword = optionValues["--source-credentials-password"].ToString();
			}
			if (optionValues.ContainsKey("--skip-old-versions"))
			{
				skipOldVersions = true;
			}
			if (optionValues.ContainsKey("--verbose"))
			{
				beVerbose = true;
			}
			if (optionValues.ContainsKey("--force-versioning"))
			{
				forceVersioning = true;
			}
			if (optionValues.ContainsKey("--versions-use-us-dates"))
			{
				versionsUseUSDates = true;
			}
			if (optionValues.ContainsKey("--prefer-folder-metadata"))
			{
				preferFolderMetadata = true;
			}
			if (optionValues.ContainsKey("--only-add-new-files-in-doclibs"))
			{
				onlyAddNewFilesInDoclibs = true;
			}
			if (optionValues.ContainsKey("--single-list"))
			{
				string[] listarg = { raw_source, raw_dest_site, raw_dest_path };
				lists.Add(listarg);
				singleList = true;
			}
			else if (optionValues.ContainsKey("--lists-file"))
			{
				lists = GetLists(optionValues["--lists-file"].ToString(), listsSeparator);
			}
			else
			{
				lists = MBSPSiteMap.GetSiteLists(raw_source, raw_dest_site);
			}
			return true;
		}


		static bool HandleArgs(string[] args, out string source, out string dest_site, out string dest_path)
		{
			source = "";
			dest_site = "";
			dest_path = "";
			int index = 0;

			if (args.Length < 1)
			{
				printUsage();
				return false;
			}

			string arg = "";
			string val = "";
			for (int i = 0; i < args.Length; i++)
			{
				arg = args[i];
				val = "true";
				if (arg.StartsWith("--"))
				{
					if (arg.Contains("="))
					{
						val = arg.Substring(arg.IndexOf("=") + 1);
						arg = arg.Substring(0, arg.IndexOf("="));
					}
					ICollection keys = options.Keys;
					foreach (object thisKey in keys)
					{
						if (arg.Equals(thisKey.ToString()))
						{
							optionValues.Add(arg, val);
						}
					}
				}
				else
				{
					index = i;
					break;
				}
			}

			if (index <= args.Length)
			{
				source = args[index];
				source = lowercaseURLHostname(source);
				if (source.EndsWith("/"))
				{
					source = source.Substring(0, source.Length - 1);
				}
			}
			if (args.Length > (index + 1))
			{
				dest_site = args[++index];
				dest_site = lowercaseURLHostname(dest_site);
			}
			if (args.Length > (index + 1))
			{
				dest_path = args[++index];
			}
			return true;
		}


		static Hashtable GetReplacements(string filePath, string separator)
		{
			Hashtable r = new Hashtable();
			if (filePath.Length > 0)
			{
				string line;
				StreamReader sr = File.OpenText(filePath);
				while ((line = sr.ReadLine()) != null)
				{
					line.Trim();
					if (! (line.StartsWith("#") || line.Length == 0))
					{
						string lookFor = line.Substring(0, line.IndexOf(separator));
						string replaceWith = line.Substring(line.IndexOf(separator) + separator.Length);
						r.Add(lookFor, replaceWith);
					}
				}
				sr.Close();
			}
			return r;
		}


		static ArrayList GetLists(string filePath, string separator)
		{
			ArrayList r = new ArrayList();
			if (filePath.Length > 0)
			{
				string line;
				StreamReader sr = File.OpenText(filePath);
				while ((line = sr.ReadLine()) != null)
				{
					line.Trim();
					if (!(line.StartsWith("#") || line.Length == 0))
					{
						string source = line.Substring(0, line.IndexOf(separator));
						source = lowercaseURLHostname(source);
						string dest_site = line.Substring(line.IndexOf(separator) + separator.Length);
						dest_site = lowercaseURLHostname(dest_site);
						string dest_path = "";
						if (dest_site.IndexOf(separator) > -1)
						{
							dest_path = dest_site.Substring(dest_site.IndexOf(separator) + separator.Length);
							dest_site = dest_site.Substring(0, dest_site.IndexOf(separator));
						}
						string[] bits = { source, dest_site, dest_path };
						r.Add(bits);
					}
				}
				sr.Close();
			}
			return r;
		}


		public static System.Net.ICredentials getSourceCredentials()
		{
			System.Net.ICredentials creds;
			if ((sourceCredentialsDomain.Length > 0)
				|| (sourceCredentialsUsername.Length > 0)
				|| (sourceCredentialsPassword.Length > 0))
			{
				creds = new System.Net.NetworkCredential(Program.sourceCredentialsUsername, Program.sourceCredentialsPassword, Program.sourceCredentialsDomain);
			}
			else
			{
				creds = System.Net.CredentialCache.DefaultCredentials;
			}
			return creds;
		}


		// Uppercase letters in a host name case major issues in some Windows dlls.
		// In the source hostname we see stack overflows, in the destination we
		// see duplicate child sites created with the same name as the parent.
		public static string lowercaseURLHostname(string url)
		{
			int startOfFQDN = url.IndexOf("//") + 2;
			int endOfFQDN = url.IndexOf("/", startOfFQDN);
			string urlPath = "";
			if (endOfFQDN.Equals(-1))
			{
				endOfFQDN = url.Length;
			}
			else
			{
				urlPath = url.Substring(endOfFQDN, url.Length - endOfFQDN);
			}
			string urlBase = url.Substring(0, endOfFQDN);
			urlBase = urlBase.ToLower();
			string newUrl = urlBase + urlPath;
			return newUrl;
		}


		public static void printUsage()
		{
			Console.WriteLine("Usage: splistcp [options] <source url> <destination site id> [<folder path>]");
			ICollection keys = options.Keys;
			Console.WriteLine("");
			Console.WriteLine("Options:");
			foreach (object thisKey in keys)
			{
				Console.WriteLine("   " + thisKey.ToString() + ": " + options[thisKey].ToString());
			}
			Console.WriteLine("Examples:");
			Console.WriteLine("splistcp http://source.server/ http://destination.server/");
			Console.WriteLine("splistcp --single-list http://source.server/path/list http://destination.server/path/list");
			Console.WriteLine("splistcp --single-list \"http://source.server/site/doclib/Forms/AllItems.aspx\" \"http://destination.server/newsite/doclib\" \"Top Parent Folder/folder2\"");
			Console.WriteLine("splistcp --single-list --verbose --create-site-from-template=\"MyTemplate.stp\" \"http://source.server/site/List Name\" \"http://destination.server/newsite/New List\"");
			Console.WriteLine("");
			Console.WriteLine("Note: this tool must be run on the destination SharePoint 2007 server, using an account that has read access to the remote site, and full control on the destination site. May also require changes on destination server as per http://support.microsoft.com/kb/896861.");
		}
	}
}
