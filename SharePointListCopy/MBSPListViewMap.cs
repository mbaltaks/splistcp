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
using System.Xml;
using Microsoft.SharePoint;


namespace SharePointListCopy
{
	class MBSPListViewMap
	{
		public MBSPListViewMap(MBSPListMap aListMap)
		{
			SPListTemplateType listType = aListMap.GetSourceListType();
			SharePointViewsWebService.Views viewService = new SharePointViewsWebService.Views();
			viewService.Url = aListMap.GetSourceSiteURL() + "/_vti_bin/Views.asmx";
			viewService.Credentials = Program.getSourceCredentials();
			XmlNode allViews = viewService.GetViewCollection(aListMap.GetSourceListName());
			foreach (XmlElement s in allViews)
			{
				if (Program.beVerbose)
				{
					Console.WriteLine("");
				}
				if (listType.Equals(SPListTemplateType.DocumentLibrary)
					&& (s.Attributes["DisplayName"].Value.ToString().Equals("Explorer View")
					|| s.Attributes["DisplayName"].Value.ToString().Equals("All Documents")))
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.PictureLibrary)
					&& (s.Attributes["DisplayName"].Value.ToString().Equals("Explorer View")
					|| s.Attributes["DisplayName"].Value.ToString().Equals("All Pictures")))
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.PictureLibrary)
					&& s.Attributes["DisplayName"].Value.ToString().Equals("Slide show view"))
				{
					continue;
				}
				if (s.Attributes["DisplayName"].Value.ToString().Length < 1)
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.Events)
					&& (s.Attributes["DisplayName"].Value.ToString().Equals("Calendar")
					|| s.Attributes["DisplayName"].Value.ToString().Equals("Current Events")
					|| s.Attributes["DisplayName"].Value.ToString().Equals("All Events")))
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.DiscussionBoard)
					&& (s.Attributes["DisplayName"].Value.ToString().Equals("Threaded")))
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.Contacts)
					&& (s.Attributes["DisplayName"].Value.ToString().Equals("All Contacts")))
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.Announcements)
					&& (s.Attributes["DisplayName"].Value.ToString().Equals("All Items")))
				{
					continue;
				}
				if (listType.Equals(SPListTemplateType.Survey))
				{
					// There are only three views, and they can't be changed or added to.
					continue;
				}
				if (Program.beVerbose)
				{
					Console.WriteLine("Copying View " + s.Attributes["DisplayName"].Value.ToString());
				}
				XmlNode dv = s.Attributes.GetNamedItem("DefaultView");
				bool defaultView = false;
				if (dv != null)
				{
					defaultView = System.Convert.ToBoolean(dv.Value.ToString());
				}

				XmlNode v = viewService.GetView(aListMap.GetSourceListName(), s.Attributes["Name"].Value.ToString());
				/*foreach (XmlAttribute attr in v.Attributes)
				{
					Console.WriteLine("");
					Console.WriteLine(attr.Name.ToString());
					Console.WriteLine(attr.Value.ToString());
				}*/
				/*foreach (XmlElement el in v.ChildNodes)
				{
					Console.WriteLine("");
					Console.WriteLine(el.OuterXml);
				}*/
				string rows = v.ChildNodes[2].ChildNodes[0].Value.ToString();
				uint rowCount = System.Convert.ToUInt32(rows, 10);
				
				SPViewCollection views = aListMap.destList.Views;
				System.Collections.Specialized.StringCollection viewFields = new System.Collections.Specialized.StringCollection();
				foreach (XmlElement el in v.ChildNodes[1])
				{
					string sourceInternalName = el.Attributes["Name"].Value.ToString();
					string displayName = aListMap.listFields[sourceInternalName].ToString();
					string newInternalName = sourceInternalName;
					if (aListMap.reverseNewListFields.ContainsKey(displayName))
					{
						newInternalName = aListMap.reverseNewListFields[displayName].ToString();
					}
					viewFields.Add(newInternalName);
				}
				string query = v.ChildNodes[0].InnerXml;
				string name = s.Attributes["DisplayName"].Value.ToString();
				bool paged = System.Convert.ToBoolean(v.ChildNodes[2].Attributes["Paged"].Value.ToString());
				SPViewCollection.SPViewType type = GetViewTypeFromTypeCode(s.Attributes["Type"].Value.ToString());
				views.Add(name, viewFields, query, rowCount, paged, defaultView, type, false);
			}
		}


		public SPViewCollection.SPViewType GetViewTypeFromTypeCode(string typeCode)
		{
			switch (typeCode)
			{
				case "HTML":
					return SPViewCollection.SPViewType.Html;
				case "GRID":
					return SPViewCollection.SPViewType.Grid;
			}
			return SPViewCollection.SPViewType.Html;
		}
	}
}
