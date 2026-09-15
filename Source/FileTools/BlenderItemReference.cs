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
	//*	BlenderItemReferenceCollection																					*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of BlenderItemReferenceItem Items.
	/// </summary>
	public class BlenderItemReferenceCollection : List<BlenderItemReferenceItem>
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
	//*	BlenderItemReferenceItem																								*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// individual reference information about the item's last usage.
	/// </summary>
	public class BlenderItemReferenceItem
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
		//*	ActionType																														*
		//*-----------------------------------------------------------------------*
		private BlenderSheetActionTypeEnum mActionType =
			BlenderSheetActionTypeEnum.None;
		/// <summary>
		/// Get/Set the action type to run on this instance.
		/// </summary>
		public BlenderSheetActionTypeEnum ActionType
		{
			get { return mActionType; }
			set { mActionType = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Comment																																*
		//*-----------------------------------------------------------------------*
		private string mComment = "";
		/// <summary>
		/// Get/Set the current comment on this instance.
		/// </summary>
		public string Comment
		{
			get { return mComment; }
			set { mComment = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	FrameIndex																														*
		//*-----------------------------------------------------------------------*
		private string mFrameIndex = "";
		/// <summary>
		/// Get/Set the frame index of this reference.
		/// </summary>
		public string FrameIndex
		{
			get { return mFrameIndex; }
			set { mFrameIndex = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ObjectName																														*
		//*-----------------------------------------------------------------------*
		private string mObjectName = "";
		/// <summary>
		/// Get/Set the name of the object.
		/// </summary>
		public string ObjectName
		{
			get { return mObjectName; }
			set { mObjectName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	RotateX																																*
		//*-----------------------------------------------------------------------*
		private string mRotateX = "";
		/// <summary>
		/// Get/Set the current rotation on the X axis.
		/// </summary>
		public string RotateX
		{
			get { return mRotateX; }
			set { mRotateX = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	RotateY																																*
		//*-----------------------------------------------------------------------*
		private string mRotateY = "";
		/// <summary>
		/// Get/Set the current rotation on the Y axis.
		/// </summary>
		public string RotateY
		{
			get { return mRotateY; }
			set { mRotateY = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	RotateZ																																*
		//*-----------------------------------------------------------------------*
		private string mRotateZ = "";
		/// <summary>
		/// Get/Set the current rotation on the Z axis.
		/// </summary>
		public string RotateZ
		{
			get { return mRotateZ; }
			set { mRotateZ = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ScaleX																																*
		//*-----------------------------------------------------------------------*
		private string mScaleX = "";
		/// <summary>
		/// Get/Set the current scale on the X axis.
		/// </summary>
		public string ScaleX
		{
			get { return mScaleX; }
			set { mScaleX = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ScaleY																																*
		//*-----------------------------------------------------------------------*
		private string mScaleY = "";
		/// <summary>
		/// Get/Set the current scale on the Y axis.
		/// </summary>
		public string ScaleY
		{
			get { return mScaleY; }
			set { mScaleY = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ScaleZ																																*
		//*-----------------------------------------------------------------------*
		private string mScaleZ = "";
		/// <summary>
		/// Get/Set the current scale on the Z axis.
		/// </summary>
		public string ScaleZ
		{
			get { return mScaleZ; }
			set { mScaleZ = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	TranslateX																														*
		//*-----------------------------------------------------------------------*
		private string mTranslateX = "";
		/// <summary>
		/// Get/Set the current position on the X axis.
		/// </summary>
		public string TranslateX
		{
			get { return mTranslateX; }
			set { mTranslateX = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	TranslateY																														*
		//*-----------------------------------------------------------------------*
		private string mTranslateY = "";
		/// <summary>
		/// Get/Set the current position on the Y axis.
		/// </summary>
		public string TranslateY
		{
			get { return mTranslateY; }
			set { mTranslateY = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	TranslateZ																														*
		//*-----------------------------------------------------------------------*
		private string mTranslateZ = "";
		/// <summary>
		/// Get/Set the current position on the Z axis.
		/// </summary>
		public string TranslateZ
		{
			get { return mTranslateZ; }
			set { mTranslateZ = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Value																																	*
		//*-----------------------------------------------------------------------*
		private string mValue = "";
		/// <summary>
		/// Get/Set the value or parameter to apply.
		/// </summary>
		public string Value
		{
			get { return mValue; }
			set { mValue = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}
