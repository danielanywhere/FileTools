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
	//*	FileFindTracker																													*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Tracking information for finding files.
	/// </summary>
	public class FileFindTracker
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
		//*	AcceptDirectories																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for
		/// <see cref="AcceptDirectories">AcceptDirectories</see>.
		/// </summary>
		private bool mAcceptDirectories = true;
		/// <summary>
		/// Get/Set a value indicating whether directories are accepted in the
		/// result.
		/// </summary>
		public bool AcceptDirectories
		{
			get { return mAcceptDirectories; }
			set { mAcceptDirectories = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	AcceptFiles																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="AcceptFiles">AcceptFiles</see>.
		/// </summary>
		private bool mAcceptFiles = true;
		/// <summary>
		/// Get/Set a value indicating whether files are accepted in the result.
		/// </summary>
		public bool AcceptFiles
		{
			get { return mAcceptFiles; }
			set { mAcceptFiles = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Directory																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Directory">Directory</see>.
		/// </summary>
		private DirectoryInfo mDirectory = null;
		/// <summary>
		/// Get/Set a reference to the directory within which the search is
		/// currently being performed.
		/// </summary>
		public DirectoryInfo Directory
		{
			get { return mDirectory; }
			set { mDirectory = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	DirectoryCount																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="DirectoryCount">DirectoryCount</see>.
		/// </summary>
		private int mDirectoryCount = 0;
		/// <summary>
		/// Get/Set the number of directories explicitly matched.
		/// </summary>
		public int DirectoryCount
		{
			get { return mDirectoryCount; }
			set { mDirectoryCount = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	EndDate																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="EndDate">EndDate</see>.
		/// </summary>
		private DateTime mEndDate = DateTime.MinValue;
		/// <summary>
		/// Get/Set the ending date at which to accept matching files.
		/// </summary>
		public DateTime EndDate
		{
			get { return mEndDate; }
			set { mEndDate = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	FileCount																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="FileCount">FileCount</see>.
		/// </summary>
		private int mFileCount = 0;
		/// <summary>
		/// Get/Set the number of matching files found.
		/// </summary>
		public int FileCount
		{
			get { return mFileCount; }
			set { mFileCount = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	FindPattern																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="FindPattern">FindPattern</see>.
		/// </summary>
		private string mFindPattern = "";
		/// <summary>
		/// Get/Set the find pattern for this instance.
		/// </summary>
		public string FindPattern
		{
			get { return mFindPattern; }
			set { mFindPattern = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Recurse																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="Recurse">Recurse</see>.
		/// </summary>
		private bool mRecurse = true;
		/// <summary>
		/// Get/Set a value indicating whether to recurse child directories during
		/// search.
		/// </summary>
		public bool Recurse
		{
			get { return mRecurse; }
			set { mRecurse = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	StartDate																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Private member for <see cref="StartDate">StartDate</see>.
		/// </summary>
		private DateTime mStartDate = DateTime.MinValue;
		/// <summary>
		/// Get/Set the starting date at which to accept matching files.
		/// </summary>
		public DateTime StartDate
		{
			get { return mStartDate; }
			set { mStartDate = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*
}
