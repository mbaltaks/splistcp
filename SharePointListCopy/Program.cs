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
 * along with bbPress; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA 02110-1301 USA
 *
 */

/*
 * Note: this tool must be run on the actual destination server.
 *
 * Known Issues:
 * - Doesn't yet handle matching lookup fields within the list, such as
 *   references to items in other lists.
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
		static Hashtable replacements = new Hashtable();
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
					listMap.Copy();
					listMap.Close();
				}
			}
			catch (Exception e)
			{
				Console.WriteLine("");
				Console.WriteLine("** Something has caused this to fail completely:");
				Console.WriteLine(e.Message);
			}
			finally
			{
				//logFile.Close();
			}
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
			if (optionValues.ContainsKey("--lists-file"))
			{
				lists = GetLists(optionValues["--lists-file"].ToString(), listsSeparator);
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
			if (optionValues.ContainsKey("--single-list"))
			{
				lists.Clear();
				string[] listarg = { raw_source, raw_dest_site, raw_dest_path };
				lists.Add(listarg);
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
			}
			if (args.Length > (index + 1))
			{
				dest_site = args[++index];
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
						string dest_site = line.Substring(line.IndexOf(separator) + separator.Length);
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
			Console.WriteLine("splistcp --single-list --avoid-duplicates \"http://source.server/site/List Name\" \"http://destination.server/newsite/New List\"");
			Console.WriteLine("");
			Console.WriteLine("Note: this tool must be run on the destination SharePoint 2007 server, using an account that has read access to the remote site, and full control on the destination site.");
		}
	}
}
