/*
 * Copyright (c). 2026 Daniel Patterson, MCSD (danielanywhere).
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 * 
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	BlenderSheetActionTypeEnum																							*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Enumeration of available Blender spreadsheet action column types.
	/// </summary>
	public enum BlenderSheetActionTypeEnum
	{
		/// <summary>
		/// No action type defined or unknown.
		/// </summary>
		None = 0,
		/// <summary>
		/// Set the text on a text object.
		/// </summary>
		SetText,
		/// <summary>
		/// Set the visibility on an object.
		/// </summary>
		SetVisibility,
		/// <summary>
		/// Set render visibility on an object.
		/// </summary>
		SetVisibilityRender,
		/// <summary>
		/// Set view visibility on an object.
		/// </summary>
		SetVisibilityView,
		/// <summary>
		/// Follow the path of another specified object.
		/// </summary>
		FollowPath
	}
	//*-------------------------------------------------------------------------*
}
