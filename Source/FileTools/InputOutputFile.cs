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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	InputOutputFileCollection																								*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of InputOutputFileItem Items.
	/// </summary>
	public class InputOutputFileCollection : List<InputOutputFileItem>
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
		//*	_Constructor																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Create a new instance of the InputOutputFileCollection Item.
		/// </summary>
		public InputOutputFileCollection()
		{
		}
		/// <summary>
		/// Create a new instance of the InputOutputFileCollection Item.
		/// </summary>
		/// <param name="size">
		/// Predefined size of the collection.
		/// </param>
		public InputOutputFileCollection(int size) : base(size)
		{

		}
		//*-----------------------------------------------------------------------*



	}
	//*-------------------------------------------------------------------------*

	//*-------------------------------------------------------------------------*
	//*	InputOutputFileItem																											*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Information about the association of two files, one in an input context,
	/// one in an output context.
	/// </summary>
	public class InputOutputFileItem
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
		//*	FirstInSequence																												*
		//*-----------------------------------------------------------------------*
		private bool mFirstInSequence = true;
		/// <summary>
		/// Get/Set a value indicating whether this is the first item in the
		/// sequence.
		/// </summary>
		public bool FirstInSequence
		{
			get { return mFirstInSequence; }
			set { mFirstInSequence = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	InputFile																															*
		//*-----------------------------------------------------------------------*
		private FileInfo mInputFile = null;
		/// <summary>
		/// Get/Set a reference to the input file.
		/// </summary>
		public FileInfo InputFile
		{
			get { return mInputFile; }
			set { mInputFile = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputFile																														*
		//*-----------------------------------------------------------------------*
		private FileInfo mOutputFile = null;
		/// <summary>
		/// Get/Set a reference to the output file.
		/// </summary>
		public FileInfo OutputFile
		{
			get { return mOutputFile; }
			set { mOutputFile = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	SourceIndex																														*
		//*-----------------------------------------------------------------------*
		private int mSourceIndex = 0;
		/// <summary>
		/// Get/Set the source index of this item.
		/// </summary>
		public int SourceIndex
		{
			get { return mSourceIndex; }
			set { mSourceIndex = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	TargetIndex																														*
		//*-----------------------------------------------------------------------*
		private int mTargetIndex = 0;
		/// <summary>
		/// Get/Set the target index of this item.
		/// </summary>
		public int TargetIndex
		{
			get { return mTargetIndex; }
			set { mTargetIndex = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}
