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
	//*	BlenderKeyframeActionTypeEnum																						*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Enumeration of available action types for setting Blender keyframes.
	/// </summary>
	public enum BlenderKeyframeActionTypeEnum
	{
		/// <summary>
		/// No action type specified or unknown.
		/// </summary>
		None = 0,
		/// <summary>
		/// Inline comments for documentation.
		/// </summary>
		Comment,
		/// <summary>
		/// Set the total frame count for the timeline.
		/// </summary>
		FrameCount,
		/// <summary>
		/// Set the current frame index.
		/// </summary>
		FrameIndex,
		/// <summary>
		/// Select the specified object.
		/// </summary>
		SelectObject,
		/// <summary>
		/// Rotate to the specified amount.
		/// </summary>
		Rotate,
		/// <summary>
		/// Translate to the specified amount.
		/// </summary>
		Translate,
		/// <summary>
		/// Scale to the specified amount.
		/// </summary>
		Scale,
		/// <summary>
		/// Set the visibility of the object.
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
		/// Set the text on the selected object.
		/// </summary>
		SetText,
		/// <summary>
		/// Set the front axis control for upcoming movements.
		/// </summary>
		SetFrontAxis,
		/// <summary>
		/// Set the previous frame index reference for elapsed time actions.
		/// </summary>
		SetPreviousFrameIndex,
		/// <summary>
		/// Follow the specified path to the current keyframe.
		/// </summary>
		FollowPath
	}
	//*-------------------------------------------------------------------------*

}
