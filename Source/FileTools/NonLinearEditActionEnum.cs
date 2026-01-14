/*
 * Copyright (c). 2000-2026 Daniel Patterson, MCSD (danielanywhere).
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
	//*	NonLinearEditActionEnum																									*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Enumeration of available non-linear editing actions.
	/// </summary>
	public enum NonLinearEditActionEnum
	{
		/// <summary>
		/// No action specified or unknown.
		/// </summary>
		None = 0,
		/// <summary>
		/// Cut a section from the source.
		/// </summary>
		Cut,
		/// <summary>
		/// Freeze the starting frame and apply that for the specified count.
		/// </summary>
		FreezeFrame,
		/// <summary>
		/// Place a mask rectangle over images in this range.
		/// </summary>
		MaskRectangle,



		/// <summary>
		/// Set all alpha values to values lower than or equal to the corresponding
		/// red-channel values found on the provided mask image.
		/// </summary>
		AlphaMask,
		/// <summary>
		/// Remove background from the specified image files. This action uses
		/// Microsoft Azure Cognitive services and requires the
		/// MSCognitiveServicesKey and MSCognitiveServicesEndpoint properties.
		/// </summary>
		RemoveBackground,
		/// <summary>
		/// Crop the working image to the provided Left, Top, Width, and Height
		/// values.
		/// </summary>
		CropImage,
		/// <summary>
		/// Make adjustments to alpha values of pixels in an image matching the
		/// specified values. Available variables are a, r, g, b.
		/// </summary>
		AlphaConditionalAdjust,
		/// <summary>
		/// Create anti-aliased edges between transparent and opaque areas.
		/// </summary>
		AntiAliasTransparency,
		/// <summary>
		/// Make adjustments to color values of pixels in an image matching the
		/// specified values. Available variables are a, r, g, b.
		/// </summary>
		ColorConditionalAdjust,
		/// <summary>
		/// Set the background color of the working image, overlaying the
		/// previous contents on the new background.
		/// </summary>
		ImageBackground,
		/// <summary>
		/// Resize the image to a new target size. This action uses the
		/// stretch technique.
		/// </summary>
		ResizeImage
	}
	//*-------------------------------------------------------------------------*

}
