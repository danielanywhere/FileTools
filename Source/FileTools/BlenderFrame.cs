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

using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	BlenderFrameCollection																									*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of BlenderFrameItem Items.
	/// </summary>
	public class BlenderFrameCollection : List<BlenderFrameItem>
	{
		//*************************************************************************
		//*	Private																																*
		//*************************************************************************
		//*************************************************************************
		//*	Protected																															*
		//*************************************************************************
		//*************************************************************************
		//*	Public																																*
		//*************************************************************************


	}
	//*-------------------------------------------------------------------------*

	//*-------------------------------------------------------------------------*
	//*	BlenderFrameItem																												*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Individual Blender frame descriptor.
	/// </summary>
	public class BlenderFrameItem
	{
		//*************************************************************************
		//*	Private																																*
		//*************************************************************************
		//*************************************************************************
		//*	Protected																															*
		//*************************************************************************
		//*************************************************************************
		//*	Public																																*
		//*************************************************************************
		//*-----------------------------------------------------------------------*
		//*	Name																																	*
		//*-----------------------------------------------------------------------*
		private BlenderKeyframeActionTypeEnum mName =
			BlenderKeyframeActionTypeEnum.None;
		/// <summary>
		/// Get/Set the action name.
		/// </summary>
		[JsonConverter(typeof(StringEnumConverter))]
		[JsonProperty(Order = 0)]
		public BlenderKeyframeActionTypeEnum Name
		{
			get { return mName; }
			set { mName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeName																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether to serialize the Name property.
		/// </summary>
		/// <returns>
		/// Value indicating whether the property should be serialized.
		/// </returns>
		public bool ShouldSerializeName()
		{
			return true;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeValue																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether to serialize the Value property.
		/// </summary>
		/// <returns>
		/// Value indicating whether the property should be serialized.
		/// </returns>
		public bool ShouldSerializeValue()
		{
			bool result = false;

			switch(mName)
			{
				case BlenderKeyframeActionTypeEnum.Comment:
				case BlenderKeyframeActionTypeEnum.FollowPath:
				case BlenderKeyframeActionTypeEnum.FrameCount:
				case BlenderKeyframeActionTypeEnum.FrameIndex:
				case BlenderKeyframeActionTypeEnum.SelectObject:
				case BlenderKeyframeActionTypeEnum.SetFrontAxis:
				case BlenderKeyframeActionTypeEnum.SetPreviousFrameIndex:
				case BlenderKeyframeActionTypeEnum.SetText:
				case BlenderKeyframeActionTypeEnum.SetVisibility:
					result = true;
					break;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeX																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether to serialize the X property.
		/// </summary>
		/// <returns>
		/// Value indicating whether the property should be serialized.
		/// </returns>
		public bool ShouldSerializeX()
		{
			bool result = false;
			switch(mName)
			{
				case BlenderKeyframeActionTypeEnum.Rotate:
				case BlenderKeyframeActionTypeEnum.Scale:
				case BlenderKeyframeActionTypeEnum.Translate:
					result = mX.Length > 0;
					break;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeY																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether to serialize the Y property.
		/// </summary>
		/// <returns>
		/// Value indicating whether the property should be serialized.
		/// </returns>
		public bool ShouldSerializeY()
		{
			bool result = false;
			switch(mName)
			{
				case BlenderKeyframeActionTypeEnum.Rotate:
				case BlenderKeyframeActionTypeEnum.Scale:
				case BlenderKeyframeActionTypeEnum.Translate:
					result = mY.Length > 0;
					break;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeZ																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether to serialize the Z property.
		/// </summary>
		/// <returns>
		/// Value indicating whether the property should be serialized.
		/// </returns>
		public bool ShouldSerializeZ()
		{
			bool result = false;
			switch(mName)
			{
				case BlenderKeyframeActionTypeEnum.Rotate:
				case BlenderKeyframeActionTypeEnum.Scale:
				case BlenderKeyframeActionTypeEnum.Translate:
					result = mZ.Length > 0;
					break;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Value																																	*
		//*-----------------------------------------------------------------------*
		private string mValue = "";
		/// <summary>
		/// Get/Set the value of the record.
		/// </summary>
		[JsonProperty(Order = 1)]
		public string Value
		{
			get { return mValue; }
			set { mValue = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	X																																			*
		//*-----------------------------------------------------------------------*
		private string mX = "";
		/// <summary>
		/// Get/Set the X coordinate of the value.
		/// </summary>
		[JsonProperty(Order = 2)]
		public string X
		{
			get { return mX; }
			set { mX = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Y																																			*
		//*-----------------------------------------------------------------------*
		private string mY = "";
		/// <summary>
		/// Get/Set the Y coordinate of the value.
		/// </summary>
		[JsonProperty(Order = 3)]
		public string Y
		{
			get { return mY; }
			set { mY = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Z																																			*
		//*-----------------------------------------------------------------------*
		private string mZ = "";
		/// <summary>
		/// Get/Set the Z coordinate of the value.
		/// </summary>
		[JsonProperty(Order = 4)]
		public string Z
		{
			get { return mZ; }
			set { mZ = value; }
		}
		//*-----------------------------------------------------------------------*


	}
	//*-------------------------------------------------------------------------*

}
