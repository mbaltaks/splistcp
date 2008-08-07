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

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("SharePointListCopy")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("Michael Baltaks")]
[assembly: AssemblyProduct("SharePointListCopy")]
[assembly: AssemblyCopyright("Copyright © Michael Baltaks 2008")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("e7a57238-74da-4b70-98ea-f0c0945d8fa0")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// Don't use this Revision number, let the version control system do that.
[assembly: AssemblyVersion("1.0.10.0")]
[assembly: AssemblyFileVersion("1.0.10.0")]

namespace SharePointListCopy
{
	class Rev
	{
		public string svnrev = "$Revision$";
		public Rev()
		{
			svnrev = svnrev.Replace("$Revision: ", "");
			svnrev = svnrev.Replace(" $", "");
		}
	}
}
