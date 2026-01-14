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
	//*	FileOptionCollection																										*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of FileOptionItem Items.
	/// </summary>
	public class FileOptionCollection : List<FileOptionItem>
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
		//*	Add																																		*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Parse and add an option from its text value.
		/// </summary>
		/// <param name="optionText">
		/// Text to parse.
		/// </param>
		/// <returns>
		/// Newly created and added option.
		/// </returns>
		public FileOptionItem Add(string optionText)
		{
			char[] comma = new char[] { ',' };
			string[] parts = null;
			FileOptionItem result = new FileOptionItem();


			if(optionText?.Length > 0)
			{
				//	Text has been provided.
				parts = optionText.Split(comma,
					StringSplitOptions.RemoveEmptyEntries |
					StringSplitOptions.TrimEntries);
				if(parts.Length > 0)
				{
					result.Name = parts[0];
					if(parts.Length > 1)
					{
						result.Value = parts[1];
					}
				}
			}
			this.Add(result);
			return result;
		}
		//*-----------------------------------------------------------------------*


	}
	//*-------------------------------------------------------------------------*

	//*-------------------------------------------------------------------------*
	//*	FileOptionItem																													*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Individual option for the current action.
	/// </summary>
	public class FileOptionItem
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
		private string mName = "";
		/// <summary>
		/// Get/Set the name of the option.
		/// </summary>
		public string Name
		{
			get { return mName; }
			set { mName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Value																																	*
		//*-----------------------------------------------------------------------*
		private string mValue = "";
		/// <summary>
		/// Get/Set the optional value of the option.
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
