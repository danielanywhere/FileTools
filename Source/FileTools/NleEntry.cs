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
using System.Data;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json;

using static FileTools.FileToolsUtil;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	NleEntryCollection																											*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of NleEntryItem Items.
	/// </summary>
	public class NleEntryCollection : List<NleEntryItem>
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
		//*	FromDataTable																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a new NleEntryCollection, constructed from a data table.
		/// </summary>
		/// <param name="table">
		/// Reference to the data table to inspect.
		/// </param>
		/// <returns>
		/// Reference to a newly created and filled NLE entry collection, if
		/// successful. Otherwise, an empty NLE collection.
		/// </returns>
		public static NleEntryCollection FromDataTable(DataTable table)
		{
			NonLinearEditActionEnum action = NonLinearEditActionEnum.None;
			DataColumn colAction = null;
			DataColumn colAssignment = null;
			DataColumn colColor = null;
			DataColumn colCondition = null;
			DataColumn colCount = null;
			DataColumn colEnd = null;
			DataColumn colHeight = null;
			DataColumn colProperties = null;
			DataColumn colRemarks = null;
			DataColumn colStart = null;
			DataColumn colWidth = null;
			DataColumn colX = null;
			DataColumn colY = null;
			NleEntryItem item = null;
			NleEntryCollection result = new NleEntryCollection();
			string text = "";

			if(table != null && table.Rows.Count > 0)
			{
				colAction = GetColumn(table, "Action");
				colAssignment = GetColumn(table, "Assignment");
				colColor = GetColumn(table, "Color");
				colCondition = GetColumn(table, "Condition");
				colCount = GetColumn(table, "Count");
				colEnd = GetColumn(table, "End");
				colHeight = GetColumn(table, "Height");
				colProperties = GetColumn(table, "Properties");
				colRemarks = GetColumn(table, "Remarks");
				colStart = GetColumn(table, "Start");
				colWidth = GetColumn(table, "Width");
				colX = GetColumn(table, "X");
				colY = GetColumn(table, "Y");
				if(colAction != null)
				{
					foreach(DataRow rowItem in table.Rows)
					{
						action = NonLinearEditActionEnum.None;
						Enum.TryParse<NonLinearEditActionEnum>(
							rowItem.Field<string>(colAction), true, out action);
						item = new NleEntryItem()
						{
							Action = action
						};
						if(colAssignment != null)
						{
							text = rowItem.Field<string>(colAssignment);
							if(text?.Length > 0)
							{
								item.Assignment = text;
							}
						}
						if(colColor != null)
						{
							text = rowItem.Field<string>(colColor);
							if(text?.Length > 0)
							{
								item.Color = text;
							}
						}
						if(colCondition != null)
						{
							text = rowItem.Field<string>(colCondition);
							if(text?.Length > 0)
							{
								item.Condition = text;
							}
						}
						if(colCount != null)
						{
							text = rowItem.Field<string>(colCount);
							if(text?.Length > 0)
							{
								item.Count = ToInt(text);
							}
						}
						if(colEnd != null)
						{
							text = rowItem.Field<string>(colEnd);
							if(text?.Length > 0)
							{
								item.End = ToInt(text);
							}
						}
						if(colHeight != null)
						{
							text = rowItem.Field<string>(colHeight);
							if(text?.Length > 0)
							{
								item.Height = ToInt(text);
							}
						}
						if(colProperties != null)
						{
							text = rowItem.Field<string>(colProperties);
							if(text?.Length > 0)
							{
								item.Properties = NameValueCollection.ParseSemi(text);
							}
						}
						if(colRemarks != null)
						{
							text = rowItem.Field<string>(colRemarks);
							if(text?.Length > 0)
							{
								item.Remarks = text;
							}
						}
						if(colStart != null)
						{
							text = rowItem.Field<string>(colStart);
							if(text?.Length > 0)
							{
								item.Start = ToInt(text);
							}
						}
						if(colWidth != null)
						{
							text = rowItem.Field<string>(colWidth);
							if(text?.Length > 0)
							{
								item.Width = ToInt(text);
							}
						}
						if(colX != null)
						{
							text = rowItem.Field<string>(colX);
							if(text?.Length > 0)
							{
								item.X = ToInt(text);
							}
						}
						if(colY != null)
						{
							text = rowItem.Field<string>(colY);
							if(text?.Length > 0)
							{
								item.Y = ToInt(text);
							}
						}
						result.Add(item);
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetCutRanges																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the collection of source cut ranges found in the NLE edit
		/// collection.
		/// </summary>
		/// <returns>
		/// Cut ranges to apply to the source numbering.
		/// </returns>
		public IntRangeCollection GetCutRanges()
		{
			IntRangeItem range = null;
			IntRangeCollection result = new IntRangeCollection();
			int temp = 0;

			foreach(NleEntryItem entryItem in this)
			{
				if(entryItem.Action == NonLinearEditActionEnum.Cut &&
					!entryItem.Mute)
				{
					range = new IntRangeItem();
					range.Start = entryItem.Start;
					if(entryItem.End > 0)
					{
						range.End = entryItem.End;
					}
					else if(entryItem.Count > 0)
					{
						//	A count was specified.
						//	Count includes the first item.
						range.End = range.Start + entryItem.Count - 1;
					}
					else
					{
						//	Neither the end nor the count were specified.
						range.End = int.MaxValue;
					}
					if(range.Start > range.End)
					{
						//	Start and end values were reversed.
						temp = range.Start;
						range.Start = range.End;
						range.End = temp;
					}
					result.Add(range);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*



	}
	//*-------------------------------------------------------------------------*

	//*-------------------------------------------------------------------------*
	//*	NleEntryItem																														*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Information about an individual NLE entry.
	/// </summary>
	public class NleEntryItem
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
		//*	Action																																*
		//*-----------------------------------------------------------------------*
		private NonLinearEditActionEnum mAction = NonLinearEditActionEnum.None;
		/// <summary>
		/// Get/Set the action to execute on this step.
		/// </summary>
		[JsonProperty(Order = 1)]
		public NonLinearEditActionEnum Action
		{
			get { return mAction; }
			set { mAction = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Assignment																														*
		//*-----------------------------------------------------------------------*
		private string mAssignment = "";
		/// <summary>
		/// Get/Set the conditional assignment associated with this entry.
		/// </summary>
		[JsonProperty(Order = 8)]
		public string Assignment
		{
			get { return mAssignment; }
			set { mAssignment = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Color																																	*
		//*-----------------------------------------------------------------------*
		private string mColor = "";
		/// <summary>
		/// Get/Set the color associated with this entry.
		/// </summary>
		[JsonProperty(Order = 8)]
		public string Color
		{
			get { return mColor; }
			set { mColor = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Condition																															*
		//*-----------------------------------------------------------------------*
		private string mCondition = "";
		/// <summary>
		/// Get/Set the condition associated with this entry.
		/// </summary>
		[JsonProperty(Order = 8)]
		public string Condition
		{
			get { return mCondition; }
			set { mCondition = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Count																																	*
		//*-----------------------------------------------------------------------*
		private int mCount = 0;
		/// <summary>
		/// Get/Set a count value.
		/// </summary>
		[JsonProperty(Order = 3)]
		public int Count
		{
			get { return mCount; }
			set { mCount = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	End																																		*
		//*-----------------------------------------------------------------------*
		private int mEnd = 0;
		/// <summary>
		/// Get/Set the end frame.
		/// </summary>
		[JsonProperty(Order = 2)]
		public int End
		{
			get { return mEnd; }
			set { mEnd = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetEnd																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the final end frame of this entry, whether the Count or End
		/// properties were used.
		/// </summary>
		/// <param name="entry">
		/// Reference to the entry to check.
		/// </param>
		/// <returns>
		/// </returns>
		public static int GetEnd(NleEntryItem entry)
		{
			int result = 0;

			if(entry != null)
			{
				if(entry.mEnd > 0)
				{
					result = entry.mEnd;
				}
				else if(entry.mCount > 0)
				{
					result = entry.mStart + entry.mCount - 1;
				}
				else
				{
					result = int.MaxValue;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Height																																*
		//*-----------------------------------------------------------------------*
		private int mHeight = 0;
		/// <summary>
		/// Get/Set the height associated with this entry.
		/// </summary>
		[JsonProperty(Order = 7)]
		public int Height
		{
			get { return mHeight; }
			set { mHeight = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Mute																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Get/Set a value indicating whether this entry is muted.
		/// </summary>
		[JsonIgnore]
		public bool Mute
		{
			get
			{
				return mProperties.Exists(x =>
					x.Name.ToLower() == "mute" && x.Value.ToLower() == "true");
			}
			set
			{
				NameValueItem property =
					mProperties.FirstOrDefault(x => x.Name.ToLower() == "mute");

				if(property == null)
				{
					property = new NameValueItem()
					{
						Name = "Mute"
					};
					mProperties.Add(property);
				}
				property.Value = value.ToString();
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Properties																														*
		//*-----------------------------------------------------------------------*
		private NameValueCollection mProperties = new NameValueCollection();
		/// <summary>
		/// Get/Set the properties associated with this entry.
		/// </summary>
		[JsonProperty(Order = 7)]
		public NameValueCollection Properties
		{
			get { return mProperties; }
			set { mProperties = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Remarks																																*
		//*-----------------------------------------------------------------------*
		private string mRemarks = "";
		/// <summary>
		/// Get/Set the remarks associated with this entry.
		/// </summary>
		[JsonProperty(Order = 8)]
		public string Remarks
		{
			get { return mRemarks; }
			set { mRemarks = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Start																																	*
		//*-----------------------------------------------------------------------*
		private int mStart = 0;
		/// <summary>
		/// Get/Set the source start frame.
		/// </summary>
		[JsonProperty(Order = 0)]		
		public int Start
		{
			get { return mStart; }
			set { mStart = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Width																																	*
		//*-----------------------------------------------------------------------*
		private int mWidth = 0;
		/// <summary>
		/// Get/Set the width associated with this entry.
		/// </summary>
		[JsonProperty(Order = 6)]
		public int Width
		{
			get { return mWidth; }
			set { mWidth = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	X																																			*
		//*-----------------------------------------------------------------------*
		private int mX = 0;
		/// <summary>
		/// Get/Set the X coordinate for this entry.
		/// </summary>
		[JsonProperty(Order = 4)]
		public int X
		{
			get { return mX; }
			set { mX = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Y																																			*
		//*-----------------------------------------------------------------------*
		private int mY = 0;
		/// <summary>
		/// Get/Set the Y coordinate for this entry.
		/// </summary>
		[JsonProperty(Order = 5)]
		public int Y
		{
			get { return mY; }
			set { mY = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}
