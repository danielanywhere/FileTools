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
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Net.Http;
using System.Net.Mime;
using System.Numerics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Wordprocessing;

using Flee.PublicTypes;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

using dotExcel;
using Html;

using static FileTools.FileToolsUtil;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	FileActionCollection																										*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Collection of FileActionItem Items.
	/// </summary>
	public class FileActionCollection : List<FileActionItem>
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
		//*	GetBase																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the base number or filename pattern of the source or target
		/// files, depending upon the action.
		/// </summary>
		public string GetBase()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.Base;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	GetBytes																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the closest ancestor instance of the binary byte buffer for this
		/// instance.
		/// </summary>
		public byte[] GetBytes()
		{
			byte[] result = new byte[0];

			if(mParent != null)
			{
				result = mParent.Bytes;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetCount																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the count property from the parent entity.
		/// </summary>
		/// <returns>
		/// The value of the count property in the parent entity, if found.
		/// Otherwise, 0.
		/// </returns>
		public float GetCount()
		{
			float result = 0f;

			if(mParent != null)
			{
				result = mParent.Count;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetCurrentFile																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the CurrentFile property from the parent entity.
		/// </summary>
		/// <returns>
		/// The value of the CurrentFile property in the parent entity, if found.
		/// Otherwise, null.
		/// </returns>
		public FileInfo GetCurrentFile()
		{
			FileInfo result = null;

			if(mParent != null)
			{
				result = mParent.CurrentFile;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetDateTimeValue																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the DateTimeValue property of the parent entity.
		/// </summary>
		/// <returns>
		/// The DateTimeValue property value of the parent entity, if found.
		/// Otherwise, DateTime.MinValue.
		/// </returns>
		public DateTime GetDateTimeValue()
		{
			DateTime result = DateTime.MinValue;

			if(mParent != null)
			{
				result = mParent.DateTimeValue;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetDigits																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the Digits property from the parent entity.
		/// </summary>
		/// <returns>
		/// The value of the Digits property on the parent entity, if found.
		/// Otherwise, 0.
		/// </returns>
		public int GetDigits()
		{
			int result = 0;

			if(mParent != null)
			{
				result = mParent.Digits;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* GetInputDir																														*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Return the InputDir property of the parent entity.
		///// </summary>
		///// <returns>
		///// Reference to the InputDir property on the parent entity, if found.
		///// Otherwise, null.
		///// </returns>
		//public DirectoryInfo GetInputDir()
		//{
		//	DirectoryInfo result = null;

		//	if(mParent != null)
		//	{
		//		result = mParent.InputDir;
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetInputFilename																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the InputFilename property from the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the InputFilename property on the parent entity, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public string GetInputFilename()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.InputFilename;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	GetInputFiles																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a reference to the collection of file information used as input
		/// in this session.
		/// </summary>
		/// <returns>
		/// A reference to the parent's InputFiles collection, if found. Otherwise,
		/// an empty collection.
		/// </returns>
		public List<FileInfo> GetInputFiles()
		{
			List<FileInfo> result = null;

			if(mParent != null)
			{
				result = mParent.InputFiles;
			}
			else
			{
				result = new List<FileInfo>();
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetInputFolderName																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the InputFolderName property value from the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the InputFolderName property on the parent entity, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public string GetInputFolderName()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.InputFolderName;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	GetInputNames																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Get a reference to the list of filenames or foldernames with
		/// or without wildcards. This parameter can be specified multiple times
		/// on the command line with different values to load multiple input files.
		/// </summary>
		public List<string> GetInputNames()
		{
			List<string> result = null;

			if(mParent != null)
			{
				//	If this item is not overridden, then default to the parent.
				result = mParent.InputNames;
			}
			else
			{
				result = new List<string>();
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetOptionByName																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the option specified by name from a parent entity.
		/// </summary>
		/// <param name="collection">
		/// Reference to the collection for which the option is being found.
		/// </param>
		/// <param name="optionName">
		/// Name of the option to retrieve.
		/// </param>
		/// <returns>
		/// Reference to the specified option, if found. Otherwise, null.
		/// </returns>
		public static FileOptionItem GetOptionByName(
			FileActionCollection collection, string optionName)
		{
			FileOptionItem result = null;

			if(collection != null && collection.mParent != null)
			{
				result =
					FileActionItem.GetOptionByName(collection.mParent, optionName);
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* GetOutputDir																													*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Return the value of the OutputDir property on the parent entity.
		///// </summary>
		///// <returns>
		///// Reference to the OutputDir property on the parent entity, if found.
		///// Otherwise, null.
		///// </returns>
		//public DirectoryInfo GetOutputDir()
		//{
		//	DirectoryInfo result = null;

		//	if(mParent != null)
		//	{
		//		result = mParent.OutputDir;
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetOutputFilename																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the OutputFilename property on the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the OutputFilename property on the parent entity, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public string GetOutputFilename()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.OutputFilename;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetOutputFolderName																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the OutputFolderName property on the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the OutputFolderName property on the parent entity, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public string GetOutputFolderName()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.OutputFolderName;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetOutputName																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the OutputName property on the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the OutputName property on the parent entity, if found.
		/// Otherwise, and empty string.
		/// </returns>
		public string GetOutputName()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.OutputName;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetOutputType																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the OutputType property on the parent entity.
		/// </summary>
		/// <returns>
		/// The value of the OutputType property on the parent entity, if found.
		/// Otherwise, RenderFileTypeEnum.Auto.
		/// </returns>
		public RenderFileTypeEnum GetOutputType()
		{
			RenderFileTypeEnum result = RenderFileTypeEnum.Auto;

			if(mParent != null)
			{
				result = mParent.OutputType;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetPattern																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the Pattern property on the parent entity.
		/// </summary>
		/// <returns>
		/// The value of the Pattern property on the parent entity, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public string GetPattern()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.Pattern;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* GetPrefix																															*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Return the Prefix property value from the parent entity.
		///// </summary>
		///// <returns>
		///// Value of the Prefix prpoerty on the parent entity, if found. Otherwise,
		///// false.
		///// </returns>
		//public bool GetPrefix()
		//{
		//	bool result = false;

		//	if(mParent != null)
		//	{
		//		result = mParent.Prefix;
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* GetPropertyByName																											*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Return the user property specified by name from a parent entity.
		///// </summary>
		///// <param name="collection">
		///// Reference to the collection for which the property is being found.
		///// </param>
		///// <param name="propertyName">
		///// Name of the property to retrieve.
		///// </param>
		///// <returns>
		///// Value of the specified property, if found. Otherwise, an empty string.
		///// </returns>
		//public static string GetPropertyByName(
		//	FileActionCollection collection, string propertyName)
		//{
		//	string result = "";

		//	if(collection != null && collection.mParent != null)
		//	{
		//		result =
		//			FileActionItem.GetPropertyByName(collection.mParent, propertyName,
		//			false);
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetRange																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a reference to the Range property on the parent entity.
		/// </summary>
		/// <returns>
		/// Reference to the Range property on the parent entity, if found.
		/// Otherwise, null.
		/// </returns>
		public StartEndItem GetRange()
		{
			StartEndItem result = null;

			if(mParent != null)
			{
				result = mParent.Range;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetSourceFolderName																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the SourceFolderName property value from the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the SourceFolderName property on the parent entity, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public string GetSourceFolderName()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.SourceFolderName;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* GetSuffix																															*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Return the value of the Suffix property on the parent entity.
		///// </summary>
		///// <returns>
		///// Value of the suffix property on the parent entity, if found. Otherwise,
		///// false.
		///// </returns>
		//public bool GetSuffix()
		//{
		//	bool result = false;

		//	if(mParent != null)
		//	{
		//		result = mParent.Suffix;
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetText																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the value of the Text property from the parent entity.
		/// </summary>
		/// <returns>
		/// Value of the Text property on the parent entity, if found. Otherwise,
		/// an empty string.
		/// </returns>
		public string GetText()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.Text;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	GetWorkingPath																												*
		//*-----------------------------------------------------------------------*
		//private string mWorkingPath = "";
		/// <summary>
		/// Return the working path for operations in this instance.
		/// </summary>
		public string GetWorkingPath()
		{
			string result = "";

			if(mParent != null)
			{
				result = mParent.WorkingPath;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* InitializeParent																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Initialize the Parent property in all of the decendants of the
		/// specified collection.
		/// </summary>
		/// <param name="actions">
		/// Reference to a collection of actions.
		/// </param>
		public static void InitializeParent(FileActionCollection actions)
		{
			if(actions?.Count > 0)
			{
				foreach(FileActionItem actionItem in actions)
				{
					actionItem.Parent = actions;
					actionItem.Actions.Parent = actionItem;
					InitializeParent(actionItem.Actions);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Parent																																*
		//*-----------------------------------------------------------------------*
		private FileActionItem mParent = null;
		/// <summary>
		/// Get/Set a reference to the batch file to which this sequence belongs.
		/// </summary>
		[JsonIgnore]
		public FileActionItem Parent
		{
			get { return mParent; }
			set
			{
				//	NOTE: This is stupid because Newtonsoft JSON ...
				//	bypasses an overridden Add(Item) method.
				mParent = value;
				foreach(FileActionItem actionItem in this)
				{
					actionItem.Parent = this;
				}
			}
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

	//*-------------------------------------------------------------------------*
	//*	FileActionItem																													*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Individual file action.
	/// </summary>
	public class FileActionItem
	{
		//*************************************************************************
		//*	Private																																*
		//*************************************************************************
		/// <summary>
		/// Public properties of this class.
		/// </summary>
		private static List<PropertyInfo> mPublicProperties =
			new List<PropertyInfo>();
		/// <summary>
		/// Working path monitor.
		/// </summary>
		private static string mWorkingPathLast = "";

		//*-----------------------------------------------------------------------*
		//* AlphaConditionalAdjust																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Conditionally adjust the alpha level of pixels on an image.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action for which the alpha values are being
		/// adjusted.
		/// </param>
		private static void AlphaConditionalAdjust(FileActionItem item)
		{
			string assignment = "";
			int assignmentResult = 0;
			Bitmap bitmap = null;
			byte[] bytes = null;
			int colorA = 0;
			int colorB = 0;
			int colorG = 0;
			int colorR = 0;
			string condition = "";
			ConditionCollection conditions = null;
			bool conditionResult = false;
			ExpressionContext context;
			int count = 0;
			IDynamicExpression dynAssignment = null;
			IDynamicExpression dynCondition = null;
			int index = 0;
			int slotWidth = 4;

			if(item != null)
			{
				conditions = GetConditions(item);
				if(conditions.Count == 0)
				{
					//	If no condition collection was presented, then use the
					//	assignment and condition properties.
					condition = GetPropertyByName(item, "Condition");
					assignment = GetPropertyByName(item, "Assignment");
					if(assignment.Length > 0 && condition.Length > 0)
					{
						conditions.Add(new ConditionItem()
						{
							Condition = condition,
							Assignment = assignment
						});
					}
				}
				if(conditions.Count > 0)
				{
					//	Conditions have been specified.
					if(CheckElements(item,
						ActionElementEnum.Inputs |
						ActionElementEnum.OutputFoldername))
					{
						if(!item.OutputDir.Exists)
						{
							item.OutputDir.Create();
						}
						context = new ExpressionContext();
						//// Allow the expression to use all static public methods of
						//// System.Math.
						//context.Imports.AddType(typeof(Math));
						context.Variables["a"] = 0;
						context.Variables["b"] = 0;
						context.Variables["g"] = 0;
						context.Variables["r"] = 0;

						foreach(FileInfo fileInfoItem in item.InputFiles)
						{
							bitmap = (Bitmap)Bitmap.FromFile(fileInfoItem.FullName);
							switch(bitmap.PixelFormat)
							{
								case PixelFormat.Format24bppRgb:
									slotWidth = 3;
									break;
								case PixelFormat.Format32bppArgb:
									slotWidth = 4;
									break;
							}
							//	Return the pixels in BGRA.
							bytes = ImageToColorBytes(bitmap);
							count = bytes.Length;
							foreach(ConditionItem conditionItem in conditions)
							{
								dynCondition =
									context.CompileDynamic(conditionItem.Condition);
								dynAssignment =
									context.CompileDynamic(conditionItem.Assignment);
								for(index = 0; index + slotWidth < count; index += slotWidth)
								{
									colorB = bytes[index];
									colorG = bytes[index + 1];
									colorR = bytes[index + 2];
									if(slotWidth > 3)
									{
										colorA = bytes[index + 3];
									}
									else
									{
										colorA = 255;
									}
									context.Variables["a"] = colorA;
									context.Variables["b"] = colorB;
									context.Variables["g"] = colorG;
									context.Variables["r"] = colorR;

									conditionResult = (bool)dynCondition.Evaluate();
									if(conditionResult)
									{
										assignmentResult = (int)dynAssignment.Evaluate();
										if(assignmentResult < 0)
										{
											assignmentResult = 0;
										}
										else if(assignmentResult > 255)
										{
											assignmentResult = 255;
										}
										if(slotWidth > 3)
										{
											bytes[index + 3] = (byte)assignmentResult;
										}
									}
								}
							}
							bitmap = ColorBytesToImage(bytes,
								bitmap.PixelFormat, bitmap.Width, bitmap.Height);
							bitmap.Save(
								Path.Combine(item.OutputDir.FullName, fileInfoItem.Name));
							Console.WriteLine($" {fileInfoItem.Name}");
						}
					}
					else
					{
						Console.WriteLine(" Error: Both the Input and OutputFolder " +
							"parameters must be supplied.");
					}
				}
				else
				{
					Console.WriteLine(" Error: Both the Condition and " +
						"Assignment properties must be supplied.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AlphaConditionalAdjustBytes																						*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Conditionally adjust the alpha level of pixels on a binary image
		/// buffer.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action for which the alpha values are being
		/// adjusted.
		/// </param>
		private static void AlphaConditionalAdjustBytes(FileActionItem item)
		{
			string assignment = "";
			int assignmentResult = 0;
			byte[] bytes = null;
			int colorA = 0;
			int colorB = 0;
			int colorG = 0;
			int colorR = 0;
			string condition = "";
			ConditionCollection conditions = null;
			bool conditionResult = false;
			ExpressionContext context;
			int count = 0;
			IDynamicExpression dynAssignment = null;
			IDynamicExpression dynCondition = null;
			int height = 0;
			int index = 0;
			int slotWidth = 4;
			int width = 0;

			if(item != null && WorkingImage?.Bitmap != null)
			{
				conditions = GetConditions(item);
				if(conditions.Count == 0)
				{
					//	If no condition collection was presented, then use the
					//	assignment and condition properties.
					condition = GetPropertyByName(item, "Condition");
					assignment = GetPropertyByName(item, "Assignment");
					if(assignment.Length > 0 && condition.Length > 0)
					{
						conditions.Add(new ConditionItem()
						{
							Condition = condition,
							Assignment = assignment
						});
					}
				}
				if(conditions.Count > 0)
				{
					switch(WorkingImage.Bitmap.PixelFormat)
					{
						case PixelFormat.Format24bppRgb:
							slotWidth = 3;
							break;
						case PixelFormat.Format32bppArgb:
							slotWidth = 4;
							break;
					}
					width = WorkingImage.Bitmap.Width;
					height = WorkingImage.Bitmap.Height;
					bytes = ImageToColorBytes(WorkingImage.Bitmap);
					count = bytes.Length;
					context = new ExpressionContext();
					//// Allow the expression to use all static public methods of
					//// System.Math.
					//context.Imports.AddType(typeof(Math));
					context.Variables["a"] = 0;
					context.Variables["b"] = 0;
					context.Variables["g"] = 0;
					context.Variables["r"] = 0;
					foreach(ConditionItem conditionItem in conditions)
					{
						Console.WriteLine($"  {conditionItem.Condition}");
						dynCondition = context.CompileDynamic(conditionItem.Condition);
						dynAssignment = context.CompileDynamic(conditionItem.Assignment);
						for(index = 0; index + slotWidth < count; index += slotWidth)
						{
							colorB = bytes[index];
							colorG = bytes[index + 1];
							colorR = bytes[index + 2];
							if(slotWidth > 3)
							{
								colorA = bytes[index + 3];
							}
							else
							{
								colorA = 255;
							}
							context.Variables["a"] = colorA;
							context.Variables["b"] = colorB;
							context.Variables["g"] = colorG;
							context.Variables["r"] = colorR;

							conditionResult = (bool)dynCondition.Evaluate();
							if(conditionResult)
							{
								assignmentResult = (int)dynAssignment.Evaluate();
								if(assignmentResult < 0)
								{
									assignmentResult = 0;
								}
								else if(assignmentResult > 255)
								{
									assignmentResult = 255;
								}
								if(slotWidth > 3)
								{
									bytes[index + 3] = (byte)assignmentResult;
								}
							}
						}
					}
					WorkingImage.Bitmap =
						ColorBytesToImage(bytes,
						WorkingImage.Bitmap.PixelFormat, width, height);
				}
				else
				{
					Console.WriteLine(" Error: Either the Conditions collection or " +
						"both the Condition and Assignment\r\n" +
						" properties must be supplied.");
				}
			}
			else
			{
				Console.WriteLine(" Error: Please load image bytes before calling " +
					"this action.");
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AlphaMask																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Using the current working image, load the specified alpha mask image
		/// and apply its maximum red value of every pixel to the working image's
		/// alpha channel.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action representing the files to process.
		/// </param>
		private static void AlphaMask(FileActionItem item)
		{
			Bitmap bitmap = null;
			int colorOffset = 0;
			int h = 0;
			byte[] maskBytes = null;
			string maskFilename = "";
			int maskA = 0;
			int maskHeight = 0;
			int maskWidth = 0;
			int targetA = 0;
			byte[] targetBytes = null;
			int targetWidth = 0;
			int targetHeight = 0;
			int w = 0;
			int x = 0;
			int y = 0;

			if(item != null && WorkingImage != null && WorkingImage.Bitmap != null)
			{
				targetWidth = WorkingImage.Bitmap.Width;
				targetHeight = WorkingImage.Bitmap.Height;
				maskFilename = GetPropertyByName(item, "MaskFilename");
				if(maskFilename.Length > 0)
				{
					maskFilename = AbsolutePath(item.WorkingPath, maskFilename);
					bitmap = (Bitmap)Bitmap.FromFile(maskFilename);
					maskWidth = bitmap.Width;
					maskHeight = bitmap.Height;
					w = Math.Min(maskWidth, targetWidth);
					h = Math.Min(maskHeight, targetHeight);
					maskBytes = ImageToColorBytes(bitmap);
					targetBytes = ImageToColorBytes(WorkingImage.Bitmap);
					for(y = 0; y < h; y ++)
					{
						for(x = 0; x < w; x ++)
						{
							maskA = maskBytes[ColorOffset(x, y, maskWidth, 2)];
							colorOffset = ColorOffset(x, y, targetWidth, 3);
							targetA = targetBytes[colorOffset];
							if(targetA > maskA)
							{
								targetBytes[colorOffset] = (byte)maskA;
							}
						}
					}
					WorkingImage.Bitmap =
						ColorBytesToImage(targetBytes, bitmap.PixelFormat,
						targetWidth, targetHeight);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AntiAliasTransparency																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Smooth hard edges between transparent and opaque.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action representing the files to process.
		/// </param>
		private static void AntiAliasTransparency(FileActionItem item)
		{
			Bitmap bitmap = null;
			byte[] bytes = null;
			int colorA = 0;
			int colorOffset = 0;
			string dist = "";
			int distance = 3;
			int f = 0;
			int h = 0;
			int index = 0;
			int mask = 0;
			int sampleA = 0;
			int w = 0;
			int x = 0;
			int xx = 0;
			int y = 0;
			int yy = 0;

			if(item != null)
			{
				dist = GetPropertyByName(item, "Distance");
				if(dist.Length > 0)
				{
					distance = ToInt(dist);
				}
				if(distance > 0)
				{
					Console.WriteLine($" Distance: {distance}");
					if(CheckElements(item,
						ActionElementEnum.Inputs |
						ActionElementEnum.OutputFoldername))
					{
						f = 255 / distance;
						if(!item.OutputDir.Exists)
						{
							item.OutputDir.Create();
						}
						foreach(FileInfo fileInfoItem in item.InputFiles)
						{
							bitmap = (Bitmap)Bitmap.FromFile(fileInfoItem.FullName);
							//	Return the pixels in BGRA.
							bytes = ImageToColorBytes(bitmap);
							w = bitmap.Width;
							h = bitmap.Height;
							//for(y = 0; y < h; y++)
							//{
							//	for(x = 0; x < w; x++)
							//	{
							//		//	Index to current line number.
							//		//	Index to current col number.
							//		//	Index to current byte number.
							//		//	Index to alpha offset.
							//		bBorder = false;
							//		colorA = bytes[(((y * w) + x) * 4) + 3];
							//		if(colorA == 0)
							//		{
							//			//	Transparency found.
							//			xx = x - 1;
							//			if(xx > 0)
							//			{
							//				//	We can look left.
							//				sampleA = bytes[(((y * w) + xx) * 4) + 3];
							//				if(sampleA > 0)
							//				{
							//					bBorder = true;
							//					break;
							//				}
							//				yy = y - 1;
							//				if(yy > 0)
							//				{
							//					//	We can look left / up.
							//					sampleA = bytes[(((yy * w) + xx) * 4) + 3];
							//					if(sampleA > 0)
							//					{
							//						bBorder = true;
							//						break;
							//					}
							//				}
							//				yy = y + 1;
							//				if(yy < h)
							//				{
							//					//	We can look left / down.
							//					sampleA = bytes[(((yy * w) + xx) * 4) + 3];
							//					if(sampleA > 0)
							//					{
							//						bBorder = true;
							//						break;
							//					}
							//				}
							//			}
							//			if(!bBorder)
							//			{
							//				xx = x + 1;
							//				if(xx < w)
							//				{
							//					//	We can look right.
							//					sampleA = bytes[(((y * w) + xx) * 4) + 3];
							//					if(sampleA > 0)
							//					{
							//						bBorder = true;
							//						break;
							//					}
							//					yy = y - 1;
							//					if(yy > 0)
							//					{
							//						//	We can look right / up.
							//						sampleA = bytes[(((yy * w) + xx) * 4) + 3];
							//						if(sampleA > 0)
							//						{
							//							bBorder = true;
							//							break;
							//						}
							//					}
							//					yy = y + 1;
							//					if(yy < h)
							//					{
							//						//	We can look right / down.
							//						sampleA = bytes[(((yy * w) + xx) * 4) + 3];
							//						if(sampleA > 0)
							//						{
							//							bBorder = true;
							//							break;
							//						}
							//					}
							//				}
							//			}
							//			if(!bBorder)
							//			{
							//				yy = y - 1;
							//				if(yy > 0)
							//				{
							//					//	We can look up.
							//					sampleA = bytes[(((yy * w) + x) * 4) + 3];
							//					if(sampleA > 0)
							//					{
							//						bBorder = true;
							//						break;
							//					}
							//				}
							//			}
							//			if(!bBorder)
							//			{
							//				yy = y + 1;
							//				if(y + 1 < h)
							//				{
							//					//	We can look down.
							//					sampleA = bytes[(((yy * w) + x) * 4) + 3];
							//					if(sampleA > 0)
							//					{
							//						bBorder = true;
							//						break;
							//					}
							//				}
							//			}
							//		}
							//	}
							//	if(bBorder)
							//	{
							//		break;
							//	}
							//}

							//	Scan for set values and smooth in block rays.
							for(y = 0; y < h; y++)
							{
								for(x = 0; x < w; x++)
								{
									//if(x >= 936 && x <= 956 && y >= 176 && y <= 196)
									//{
									//	//	TODO: !1 - Stopped here...
									//	//	TODO: The entry 946, 186 should not have empty alpha.
									//	Console.WriteLine("AntialiasTransparency. Break here...");
									//}
									colorOffset = ColorOffset(x, y, w, 3);
									//if(colorOffset == 904867)
									//{
									//	Trace.WriteLine("AntiAliasTransparency. Break here...");
									//}
									colorA = bytes[colorOffset];
									if(colorA == 0)
									{
										xx = x - distance;
										for(index = 0; xx < x && index < distance; xx++, index++)
										{
											if(xx > -1)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(xx, y, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Right.
										xx = x + distance;
										for(index = 1; xx > x && index < distance; xx--, index++)
										{
											if(xx < w)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(xx, y, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Up.
										yy = y - distance;
										for(index = 1; yy < y && index < distance; yy++, index++)
										{
											if(yy > -1)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(x, yy, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Down.
										yy = y + distance;
										for(index = 1; yy > y && index < distance; yy--, index++)
										{
											if(yy < h)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(x, yy, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Left / Up.
										xx = x - (distance / 2);
										yy = y - (distance / 2);
										for(index = 1;
											xx < x && yy < y && index < distance;
											xx++, yy++, index++)
										{
											if(xx > -1 && yy > -1)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(xx, yy, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Right / Up.
										xx = x + (distance / 2);
										yy = y - (distance / 2);
										for(index = 1;
											xx > x && yy < y && index < distance;
											xx--, yy++, index++)
										{
											if(xx < w && yy > -1)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(xx, yy, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Left / Down.
										xx = x - (distance / 2);
										yy = y + (distance / 2);
										for(index = 1;
											xx < x && yy > y && index < distance;
											xx++, yy--, index++)
										{
											if(xx > -1 && yy < h)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(xx, yy, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
										//	Right / Down.
										xx = x + (distance / 2);
										yy = y + (distance / 2);
										for(index = 1;
											xx > x && yy > y && index < distance;
											xx--, yy--, index++)
										{
											if(xx < w && yy < h)
											{
												mask = 255 - (index * f);
												colorOffset = ColorOffset(xx, yy, w, 3);
												//if(colorOffset == 904867)
												//{
												//	Trace.WriteLine("AntiAliasTransparency. Break here...");
												//}
												sampleA = bytes[colorOffset];
												if(sampleA > mask)
												{
													//	Tone the sample down to the amount allowed in
													//	this slot.
													sampleA = mask;
													bytes[colorOffset] = (byte)sampleA;
												}
											}
										}
									}
								}
							}

							bitmap = ColorBytesToImage(bytes,
								bitmap.PixelFormat,
								bitmap.Width, bitmap.Height);
							bitmap.Save(
								Path.Combine(item.OutputDir.FullName, fileInfoItem.Name));
							Console.WriteLine($" {fileInfoItem.Name}");
						}
					}
					else
					{
						Console.WriteLine(" Error: Both the Input and OutputFolder " +
							"parameters must be supplied.");
					}
				}
				else
				{
					Console.WriteLine(" Error: Distance is 0.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AntiAliasTransparencyBytes																						*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Smooth hard edges between transparent and opaque.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action representing the files to process.
		/// </param>
		private static void AntiAliasTransparencyBytes(FileActionItem item)
		{
			//Bitmap bitmap = null;
			byte[] bytes = null;
			int colorA = 0;
			int colorOffset = 0;
			string dist = "";
			int distance = 3;
			int f = 0;
			int h = 0;
			int index = 0;
			int mask = 0;
			int sampleA = 0;
			int w = 0;
			int x = 0;
			int xx = 0;
			int y = 0;
			int yy = 0;

			if(item != null && WorkingImage?.Bitmap != null)
			{
				dist = GetPropertyByName(item, "Distance");
				if(dist.Length > 0)
				{
					distance = ToInt(dist);
				}
				if(distance > 0)
				{
					Console.WriteLine($" Distance: {distance}");

					w = WorkingImage.Bitmap.Width;
					h = WorkingImage.Bitmap.Height;
					//	Return the pixels in BGRA.
					bytes = ImageToColorBytes(WorkingImage.Bitmap);

					f = 255 / distance;

					//	Scan for set values and smooth in block rays.
					for(y = 0; y < h; y++)
					{
						for(x = 0; x < w; x++)
						{
							colorOffset = ColorOffset(x, y, w, 3);
							colorA = bytes[colorOffset];
							if(colorA == 0)
							{
								xx = x - distance;
								for(index = 0; xx < x && index < distance; xx++, index++)
								{
									if(xx > -1)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(xx, y, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Right.
								xx = x + distance;
								for(index = 1; xx > x && index < distance; xx--, index++)
								{
									if(xx < w)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(xx, y, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Up.
								yy = y - distance;
								for(index = 1; yy < y && index < distance; yy++, index++)
								{
									if(yy > -1)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(x, yy, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Down.
								yy = y + distance;
								for(index = 1; yy > y && index < distance; yy--, index++)
								{
									if(yy < h)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(x, yy, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Left / Up.
								xx = x - (distance / 2);
								yy = y - (distance / 2);
								for(index = 1;
									xx < x && yy < y && index < distance;
									xx++, yy++, index++)
								{
									if(xx > -1 && yy > -1)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(xx, yy, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Right / Up.
								xx = x + (distance / 2);
								yy = y - (distance / 2);
								for(index = 1;
									xx > x && yy < y && index < distance;
									xx--, yy++, index++)
								{
									if(xx < w && yy > -1)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(xx, yy, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Left / Down.
								xx = x - (distance / 2);
								yy = y + (distance / 2);
								for(index = 1;
									xx < x && yy > y && index < distance;
									xx++, yy--, index++)
								{
									if(xx > -1 && yy < h)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(xx, yy, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
								//	Right / Down.
								xx = x + (distance / 2);
								yy = y + (distance / 2);
								for(index = 1;
									xx > x && yy > y && index < distance;
									xx--, yy--, index++)
								{
									if(xx < w && yy < h)
									{
										mask = 255 - (index * f);
										colorOffset = ColorOffset(xx, yy, w, 3);
										sampleA = bytes[colorOffset];
										if(sampleA > mask)
										{
											//	Tone the sample down to the amount allowed in
											//	this slot.
											sampleA = mask;
											bytes[colorOffset] = (byte)sampleA;
										}
									}
								}
							}
						}
					}

					WorkingImage.Bitmap =
						ColorBytesToImage(bytes, WorkingImage.Bitmap.PixelFormat, w, h);

				}
				else
				{
					Console.WriteLine(" Error: Distance is 0.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AssignOutputFolder																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assign the output folder based upon output folder name, if present,
		/// then input name if necessary.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action for which the output folder is being
		/// resolved.
		/// </param>
		/// <returns>
		/// True if the output directory object was set. Otherwise, false.
		/// </returns>
		private static bool AssignOutputFolder(FileActionItem item)
		{
			bool result = false;

			if(item != null)
			{
				if(item.InputDir != null && item.OutputDir == null &&
					item.OutputName.Length == 0)
				{
					item.OutputDir = item.InputDir;
				}
				else if(item.OutputDir == null)
				{
					item.OutputDir = new DirectoryInfo(
							AbsolutePath(
								GetPropertyByName(item, nameof(WorkingPath)),
								GetPropertyByName(item, nameof(OutputName))));
				}
				if(item.OutputDir != null)
				{
					AssureFolder(item.OutputDir.FullName, true, " Output");
					result = true;
				}
				else
				{
					Console.WriteLine(" Error: Could not resolve output folder.");
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* BuildPathProperty																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Build an absolute path value in the provided user properties
		/// PropertyName and PropertyValue at this item's immediate parent level.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action for which a path name is being specified.
		/// </param>
		private static void BuildPathProperty(FileActionItem item)
		{
			string propertyName = "";

			if(item != null && item.mParent.Parent != null)
			{
				//	Item and local properties were supplied.
				propertyName = GetPropertyByName(item, "PropertyName");
				Console.WriteLine($" Preparing property {propertyName}");
				SetPropertyValue(item.mParent.Parent, propertyName,
					AbsolutePath(
						GetPropertyByName(item, nameof(WorkingPath)),
						GetPropertyByName(item, "PropertyValue")));
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* CheckElements																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Check all of the specified elements and return a value indicating
		/// whether the masked items were all valid.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item for which the elements are being
		/// tested.
		/// </param>
		/// <param name="element">
		/// Bitmasked action element flags to require on this action.
		/// </param>
		/// <returns>
		/// Value indicating whether the check was successful.
		/// </returns>
		/// <remarks>
		/// Error messages are printed to the console when one or more of the
		/// specified elements are not found.
		/// </remarks>
		private static bool CheckElements(FileActionItem item,
			ActionElementEnum element)
		{
			int count = 0;
			DirectoryInfo dir = null;
			FileInfo file = null;
			int index = 0;
			bool result = true;
			string workingFolder = "";

			if(item != null && element != ActionElementEnum.None)
			{
				if((element & ActionElementEnum.Action) != ActionElementEnum.None)
				{
					if(item.Action == ActionTypeEnum.None)
					{
						Console.WriteLine(" Error: No action specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Base) != ActionElementEnum.None)
				{
					if(item.Base.Length == 0)
					{
						Console.WriteLine(" Error: Base is required in this action.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Count) != ActionElementEnum.None)
				{
					if(item.Count == 0f)
					{
						Console.WriteLine(" Error: Count is required for this action.");
						result = false;
					}
				}
				if((element & ActionElementEnum.DateTimeValue) !=
					ActionElementEnum.None)
				{
					if(item.DateTimeValue == DateTime.MinValue)
					{
						Console.WriteLine(" Error: DateTime was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Digits) != ActionElementEnum.None)
				{
					if(item.Digits == 0)
					{
						Console.WriteLine(" Error: A value is required for Digits.");
						result = false;
					}
				}
				if((element & ActionElementEnum.InputFilename) !=
					ActionElementEnum.None)
				{
					//	In this version, when InputFilename is expressed, only files
					//	are specified in the InputFiles collection.
					count = item.InputFiles.Count;
					for(index = 0; index < count; index++)
					{
						file = item.InputFiles[index];
						if((file.Attributes & FileAttributes.Directory) !=
							(FileAttributes)0)
						{
							//	This item is a directory. Remove it.
							item.InputFiles.RemoveAt(index);
							index--;	//	Deindex.
							count--;	//	Decount.
						}
					}
					if(item.InputFiles.Count == 0)
					{
						Console.WriteLine(" Error: Input files were not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.InputFolderName) !=
					ActionElementEnum.None)
				{
					//	In this version, when InputFoldername is expressed, only
					//	folders are specified in the InputFiles collection.
					count = item.InputFiles.Count;
					for(index = 0; index < count; index ++)
					{
						file = item.InputFiles[index];
						if((file.Attributes & FileAttributes.Directory) ==
							(FileAttributes)0)
						{
							//	This item is a file. Remove it.
							item.InputFiles.RemoveAt(index);
							index--;  //	Deindex.
							count--;	//	Decount.
						}
					}
					if(item.InputFiles.Count > 0)
					{
						//	Input folders are present.
						item.InputDir = new DirectoryInfo(item.InputFiles[0].FullName);
					}
					if(item.InputFiles.Count == 0)
					{
						//	If no files are specified, use the working folder.
						workingFolder = GetPropertyByName(item, nameof(WorkingPath));
						if(workingFolder.Length > 0 && Directory.Exists(workingFolder))
						{
							file = new FileInfo(workingFolder);
							item.InputFiles.Add(file);
							item.InputDir = new DirectoryInfo(workingFolder);
						}
					}
					if(item.InputFiles.Count == 0)
					{
						Console.WriteLine(" Error: Input folders were not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Inputs) != ActionElementEnum.None)
				{
					//	Multiple input files.
					if(item.InputFiles.Count == 0)
					{
						Console.WriteLine(" Error: No input files specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.OutputFilename) !=
					ActionElementEnum.None)
				{
					if(item.OutputFile == null)
					{
						Console.WriteLine(" Error: Output filename was not specified.");
						result = false;
					}
					else
					{
						dir = new DirectoryInfo(
							Path.GetDirectoryName(item.OutputFile.FullName));
						if(!dir.Exists)
						{
							try
							{
								dir.Create();
							}
							catch
							{
								Console.WriteLine(
									" Error: Could not create output directory.");
							}
						}
					}
				}
				if((element & ActionElementEnum.OutputFoldername) !=
					ActionElementEnum.None)
				{
					if(item.OutputDir == null)
					{
						//	If no output folder was specified, use the working folder.
						workingFolder = GetPropertyByName(item, nameof(WorkingPath));
						if(workingFolder.Length > 0 && Directory.Exists(workingFolder))
						{
							dir = new DirectoryInfo(workingFolder);
							item.OutputDir = dir;
						}
					}
					if(item.OutputDir == null)
					{
						Console.WriteLine(" Error: Output folder name was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.OutputName) != ActionElementEnum.None)
				{
					//	In this version, output can be either a file or a folder.
					if(item.OutputDir == null && item.OutputFile == null)
					{
						Console.WriteLine(
							" Error: Output name was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Pattern) != ActionElementEnum.None)
				{
					if(item.Pattern.Length == 0)
					{
						Console.WriteLine(" Error: Pattern was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Range) != ActionElementEnum.None)
				{
					//	In this version, the range can be a single ended specification.
					if(item.Range.StartValue.Length == 0 &&
						item.Range.EndValue.Length == 0)
					{
						Console.WriteLine(" Error: Range was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.SourceFolderName) !=
					ActionElementEnum.None)
				{
					//	This version only has one source folder.
					item.SourceDir = null;
					if(item.SourceFolderName.Length > 0)
					{
						//	Source folder name has been specified.
						item.SourceDir = new DirectoryInfo(
							AbsolutePath(item.WorkingPath, item.SourceFolderName));
						if(!item.SourceDir.Exists)
						{
							//	If the folder doesn't exist, release it.
							item.SourceDir = null;
						}
					}
					if(item.SourceDir == null)
					{
						Console.WriteLine(" Error: Source folder was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.Text) != ActionElementEnum.None)
				{
					if(item.Text.Length == 0)
					{
						Console.WriteLine(" Error: Text parameter was not specified.");
						result = false;
					}
				}
				if((element & ActionElementEnum.WorkingPath) !=
					ActionElementEnum.None)
				{
					if(item.WorkingPath.Length == 0)
					{
						Console.WriteLine(" Error: Working path was not specified.");
						result = false;
					}
					else
					{
						dir = new DirectoryInfo(
							GetPropertyByName(item, nameof(WorkingPath)));
						if(!dir.Exists)
						{
							Console.WriteLine(" Error: Working path does not exist.");
							result = false;
						}
						else if((dir.Attributes & FileAttributes.Directory) !=
							FileAttributes.Directory)
						{
							Console.Write(" Error: A file was specified as the working ");
							Console.WriteLine("directory.");
							result = false;
						}
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ClearInputFiles																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Clear the local InputFiles collection for the immediate parent item.
		/// </summary>
		/// <param name="item">
		/// Reference to the action item calling for the InputFiles collection to
		/// be cleared.
		/// </param>
		private static void ClearInputFiles(FileActionItem item)
		{
			if(item != null && item.mParent != null && item.mParent.Parent != null)
			{
				item.mParent.Parent.mInputFiles.Clear();
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ColorConditionalAdjustBytes																						*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Conditionally adjust the color of pixels on a binary image
		/// buffer.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action for which the color values are being
		/// adjusted.
		/// </param>
		private static void ColorConditionalAdjustBytes(FileActionItem item)
		{
			string assignment = "";
			//int assignmentResult = 0;
			byte[] bytes = null;
			System.Drawing.Color color = System.Drawing.Color.Empty;
			int colorA = 0;
			int colorB = 0;
			int colorG = 0;
			int colorR = 0;
			string condition = "";
			ConditionCollection conditions = null;
			bool conditionResult = false;
			ExpressionContext context;
			int count = 0;
			//IDynamicExpression dynAssignment = null;
			IDynamicExpression dynCondition = null;
			int height = 0;
			int index = 0;
			int slotWidth = 4;
			int width = 0;

			if(item != null && WorkingImage?.Bitmap != null)
			{
				conditions = GetConditions(item);
				if(conditions.Count == 0)
				{
					//	If no condition collection was presented, then use the
					//	assignment and condition properties.
					condition = GetPropertyByName(item, "Condition");
					assignment = GetPropertyByName(item, "Assignment");
					if(assignment.Length > 0 && condition.Length > 0)
					{
						conditions.Add(new ConditionItem()
						{
							Condition = condition,
							Assignment = assignment
						});
					}
				}
				if(conditions.Count > 0)
				{
					switch(WorkingImage.Bitmap.PixelFormat)
					{
						case PixelFormat.Format24bppRgb:
							slotWidth = 3;
							break;
						case PixelFormat.Format32bppArgb:
							slotWidth = 4;
							break;
					}
					width = WorkingImage.Bitmap.Width;
					height = WorkingImage.Bitmap.Height;
					bytes = ImageToColorBytes(WorkingImage.Bitmap);
					count = bytes.Length;
					context = new ExpressionContext();
					//// Allow the expression to use all static public methods of
					//// System.Math.
					//context.Imports.AddType(typeof(Math));
					context.Variables["a"] = 0;
					context.Variables["b"] = 0;
					context.Variables["g"] = 0;
					context.Variables["r"] = 0;
					foreach(ConditionItem conditionItem in conditions)
					{
						Console.WriteLine($"  {conditionItem.Condition}");
						dynCondition = context.CompileDynamic(conditionItem.Condition);
						//dynAssignment = context.CompileDynamic(conditionItem.Assignment);

						for(index = 0; index + slotWidth < count; index += slotWidth)
						{
							//if(index == 498459)
							//{
							//	Console.WriteLine("Index 498459. Break here...");
							//}
							//Console.WriteLine($" Index: {index}");
							colorB = bytes[index];
							colorG = bytes[index + 1];
							colorR = bytes[index + 2];
							if(slotWidth > 3)
							{
								colorA = bytes[index + 3];
							}
							else
							{
								colorA = 255;
							}
							context.Variables["a"] = colorA;
							context.Variables["b"] = colorB;
							context.Variables["g"] = colorG;
							context.Variables["r"] = colorR;

							conditionResult = (bool)dynCondition.Evaluate();
							if(conditionResult && conditionItem.Assignment?.Length > 0 &&
								conditionItem.Assignment.StartsWith('#'))
							{
								//	Condition matched.
								//	The parsiing pattern for this is #AARRGGBB.
								color = ColorTranslator.FromHtml(conditionItem.Assignment);
								bytes[index] = color.B;
								bytes[index + 1] = color.G;
								bytes[index + 2] = color.R;
								if(slotWidth > 3)
								{
									bytes[index + 3] = color.A;
								}

								//assignmentResult = (int)dynAssignment.Evaluate();
								//if(assignmentResult < 0)
								//{
								//	assignmentResult = 0;
								//}
								//else if(assignmentResult > 255)
								//{
								//	assignmentResult = 255;
								//}
								//bytes[index + 3] = (byte)assignmentResult;
							}
						}
					}
					WorkingImage.Bitmap =
						ColorBytesToImage(bytes,
							WorkingImage.Bitmap.PixelFormat, width, height);
				}
				else
				{
					Console.WriteLine(" Error: Either the Conditions collection or " +
						"both the Condition and Assignment\r\n" +
						" properties must be supplied.");
				}
			}
			else
			{
				Console.WriteLine(" Error: Please load image bytes before calling " +
					"this action.");
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ConvertFromB64																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Convert the input file from base-64 to binary.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void ConvertFromB64(FileActionItem item)
		{
			byte[] bytes = null;
			string content = "";
			string filename = "";

			if(item != null &&
				item.InputFiles.Count > 0  &&
				item.OutputFilename?.Length > 0)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFilename | ActionElementEnum.OutputFilename))
				{
					//	Input and output filenames were both provided.
					filename = item.InputFiles[0].FullName;
					content = File.ReadAllText(filename);
					//	Remove any prefix like Data URL, etc.
					content = RightOf(content, ",");
					bytes = Convert.FromBase64String(content);
					File.WriteAllBytes(item.OutputFilename, bytes);
					Console.WriteLine($" File created: {item.OutputFile.Name}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ConvertToB64																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Convert the input file from binary to base-64.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item begin fulfilled.
		/// </param>
		private static void ConvertToB64(FileActionItem item)
		{
			bool bDataUrl = false;
			byte[] bytes = null;
			string content = "";
			string extension = "";
			string filename = "";
			FileOptionItem option = null;
			string optionValue = "";

			if(item != null &&
				item.InputFiles.Count > 0 &&
				item.OutputFilename?.Length > 0)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFilename | ActionElementEnum.OutputFilename))
				{
					option = GetOptionByName(item, "dataurl");
					if(option?.Value.Length > 0)
					{
						optionValue = option.Value;
					}
					else
					{
						optionValue = GetPropertyByName(item, "dataurl");
					}
					bDataUrl = (optionValue.ToLower() == "true");

					//	Input and output filenames were both provided.
					filename = item.mInputFiles[0].FullName;
					extension = item.mInputFiles[0].Extension;
					bytes = File.ReadAllBytes(filename);
					if(bDataUrl)
					{
						content = DataUrl.ToB64(bytes, Mime.MimeType(extension));
					}
					else
					{
						content = Convert.ToBase64String(bytes);
					}
					File.WriteAllText(item.OutputFilename, content);
					Console.WriteLine($" File created: {item.OutputFile.Name}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* CopyNumericToRange																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		///	Copy the specfied numeric frame to the specified range. Used for
		///	duplicating an individual frame to multiple frames.
		/// </summary>
		/// <param name="item">
		/// Reference to the action item being fulfilled.
		/// </param>
		private static void CopyNumericToRange(FileActionItem item)
		{
			FileInfo sourceFile = null;
			List<string> targetFilenames = null;

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFilename |
					ActionElementEnum.OutputFoldername |
					ActionElementEnum.Range))
				{
					targetFilenames = EnumerateRange(item.Range, item.Digits,
						Path.GetExtension(item.InputNames[0]));
					if(targetFilenames.Count > 0)
					{
						//	Filenames were generated.
						//	Make sure we don't cause a collision.
						if(!AnyExist(item.OutputDir.FullName, targetFilenames))
						{
							//	The range is clear for copying.
							sourceFile = item.InputFiles[0];
							foreach(string filenameItem in targetFilenames)
							{
								sourceFile.CopyTo(
									AbsolutePath(item.OutputDir.FullName, filenameItem));
								Console.WriteLine($" {filenameItem}");
							}
						}
						else
						{
							//	Target files are in the way.
							Console.Write(" Error: Files exist with the ");
							Console.WriteLine("specified target range...");
						}
					}
					else
					{
						Console.Write(" Error: Target filenames could not be ");
						Console.WriteLine("generated from the given range.");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* CopyRange																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Copy a range of numbered files to the target base offset.
		/// </summary>
		/// <param name="item">
		/// Reference to the action item being fulfilled.
		/// </param>
		private static void CopyRange(FileActionItem item)
		{
			bool bDifferentFolders = false;
			int count = 0;
			int index = 0;
			string name = "";
			FileInfo sourceFile = null;
			string sourceFilename = "";
			List<string> sourceFilenames = null;
			string targetFilename = "";
			List<string> targetFilenames = null;

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFolderName |
					ActionElementEnum.Range))
				{
					if(AssignOutputFolder(item))
					{
						bDifferentFolders =
							(item.OutputDir.FullName != item.InputDir.FullName);
						if(item.Base.Length == 0)
						{
							if(bDifferentFolders)
							{
								//	If we are outputting to a different folder, then
								//	the starting filename is the base.
								item.Base = item.Range.StartValue;
							}
							else
							{
								//	If no base was specified, then the first target item
								//	starts after the end of the source range.
								item.Base =
									IncrementFilename(item.Range.EndValue, item.Digits);
							}
						}
						if(item.Base.Length > 0 &&
							item.Range.StartValue.Contains('.') &&
							item.Range.EndValue.Contains('.'))
						{
							sourceFilenames = EnumerateRange(item.Range, item.Digits);
							if(sourceFilenames.Count > 0)
							{
								//	Source filenames were generated.
								targetFilenames = new List<string>();
								name = GetPropertyByName(item, nameof(Base));
								targetFilenames.Add(name);
								count = sourceFilenames.Count;
								for(index = 1; index < count; index++)
								{
									name = IncrementFilename(name, item.Digits);
									if(name.Length > 0)
									{
										targetFilenames.Add(name);
									}
								}
								//	Make sure we don't cause a collision.
								if(bDifferentFolders ||
									!AnyExist(item.OutputDir.FullName, targetFilenames))
								{
									//	The range is clear for copying.
									count = sourceFilenames.Count;
									for(index = 0; index < count; index++)
									{
										sourceFilename = sourceFilenames[index];
										sourceFile =
											new FileInfo(
												Path.Combine(item.InputDir.FullName, sourceFilename));
										if(sourceFile.Exists && index < targetFilenames.Count)
										{
											targetFilename = targetFilenames[index];
											sourceFile.CopyTo(
												Path.Combine(item.OutputDir.FullName, targetFilename),
												bDifferentFolders);
											Console.WriteLine($" {targetFilename}");
										}
									}
								}
								else
								{
									//	Target files are in the way.
									Console.Write(" Error: Files exist with the ");
									Console.WriteLine("specified target range...");
								}
							}
							else
							{
								Console.Write(" Error: Source filenames could not be ");
								Console.WriteLine("enumerated from the given range.");
							}
						}
						else if(item.Base.Length == 0)
						{
							Console.Write(" Error: Base was not supplied and ");
							Console.WriteLine(
								$"{item.Range.EndValue} can't be incremented.");
						}
						else
						{
							Console.Write(" Error: This action requires extension names ");
							Console.WriteLine("in the range.");
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* CropImage																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Crop the working image to the rectangle specified in the user
		/// properties Left, Top, Width, and Height.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item describing the rectangle to use on
		/// the working image.
		/// </param>
		private static void CropImage(FileActionItem item)
		{
			Bitmap bitmap = null;
			int h = 0;
			Rectangle sourceRect = Rectangle.Empty;
			Rectangle targetRect = Rectangle.Empty;
			int w = 0;
			int x = 0;
			int y = 0;

			if(item != null && WorkingImage != null)
			{
				//	The item and the working image are both present.
				x = ToInt(GetPropertyByName(item, "Left"));
				y = ToInt(GetPropertyByName(item, "Top"));
				w = ToInt(GetPropertyByName(item, "Width"));
				h = ToInt(GetPropertyByName(item, "Height"));
				if(w > 0 && h > 0)
				{
					//	Width and height were supplied.
					Console.WriteLine($"  {x}, {y}, {w}, {h}");
					sourceRect = new Rectangle(x, y, w, h);
					targetRect = new Rectangle(0, 0, w, h);
					bitmap = new Bitmap(targetRect.Width, targetRect.Height);
					using(Graphics graphics = Graphics.FromImage(bitmap))
					{
						InitializeGraphics(graphics);
						graphics.DrawImage(WorkingImage.Bitmap,
							targetRect, sourceRect, GraphicsUnit.Pixel);
					}
					WorkingImage.Bitmap = bitmap;
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* CropImageToRectangleInfoName																					*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Crop the working image to the rectangle specified in the item of
		/// the rectangle info list having Name = user property Name.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item describing the name of the
		/// rectangle to use on the working image.
		/// </param>
		private static void CropImageToRectangleInfoName(FileActionItem item)
		{
			Bitmap bitmap = null;
			string name = "";
			RectInfoItem rectInfo = null;
			Rectangle sourceRect = Rectangle.Empty;
			Rectangle targetRect = Rectangle.Empty;

			if(item != null && WorkingImage != null)
			{
				//	The item and the working image are both present.
				name = GetPropertyByName(item, "Name");
				Console.WriteLine($" {name}");
				if(name?.Length > 0)
				{
					//	Name was supplied.
					rectInfo = RectangleInfoList.FirstOrDefault(x => x.Name == name);
					if(rectInfo != null)
					{
						//	Rectangle found.
						sourceRect = new Rectangle(
							(int)rectInfo.Left, (int)rectInfo.Top,
							(int)rectInfo.Width, (int)rectInfo.Height);
						targetRect = new Rectangle(0, 0,
							sourceRect.Width, sourceRect.Height);
						bitmap = new Bitmap(targetRect.Width, targetRect.Height);
						using(Graphics graphics = Graphics.FromImage(bitmap))
						{
							InitializeGraphics(graphics);
							graphics.DrawImage(WorkingImage.Bitmap,
								targetRect, sourceRect, GraphicsUnit.Pixel);
						}
						WorkingImage.Bitmap = bitmap;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DeepCopy																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a new instance of the provided item.
		/// </summary>
		/// <param name="item">
		/// Reference to the item to be copied.
		/// </param>
		/// <returns>
		/// Reference to a completely new instance of the file action item
		/// provided by the caller.
		/// </returns>
		public static FileActionItem DeepCopy(FileActionItem item)
		{
			string content = "";
			FileActionItem result = null;

			if(item != null)
			{
				content = JsonConvert.SerializeObject(item);
				result = JsonConvert.DeserializeObject<FileActionItem>(content);
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DelDirectoryPattern																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Delete the matching directory pattern.
		/// </summary>
		/// <param name="item">
		/// Reference to the action describing the directory pattern to delete.
		/// </param>
		public static void DelDirectoryPattern(FileActionItem item)
		{
			int count = 0;
			string folderName = "";
			List<string> folderNames = null;
			string workingPath = "";

			if(item != null)
			{
				folderName = GetPropertyByName(item, "InputFolderName");
				if(folderName.Length > 0)
				{
					Console.WriteLine($" InputFolderName: {folderName}");
					workingPath = GetPropertyByName(item, nameof(WorkingPath));
					workingPath = AbsolutePath(workingPath, folderName);
					folderName = workingPath;
					workingPath = FindLastDirectoryLevel(workingPath);
					if(workingPath.Length < folderName.Length)
					{
						//	Working folder name will contain wildcards.
						folderName = folderName.Substring(workingPath.Length);
						while(folderName.StartsWith('\\') || folderName.StartsWith('/'))
						{
							folderName = folderName.Substring(1);
						}
						folderNames = ResolveWildcardFolders(workingPath, folderName);
					}
					else
					{
						//	Only one folder name is present.
						folderNames = ResolveWildcardFolders(workingPath, "");
					}
					//	The Resolve Wildcard Folders method returns only existing
					//	folders.
					if(folderNames.Count > 0)
					{
						foreach(string folderNameItem in folderNames)
						{
							if(Directory.Exists(folderNameItem))
							{
								try
								{
									Directory.Delete(folderNameItem, true);
									Console.WriteLine($" Directory deleted: {folderNameItem}");
									count++;
								}
								catch(Exception ex)
								{
									Console.WriteLine($" Could not delete {folderNameItem}." +
										$" {ex.Message}");
								}
							}
						}
						if(count > 0)
						{
							Console.WriteLine($" {count} folders deleted.");
						}
					}
					else
					{
						Console.WriteLine(" No matching folders found to delete.");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DeleteFile																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Delete the file specified in the Filename user property.
		/// </summary>
		/// <param name="item">
		/// Reference to the action describing the file to delete.
		/// </param>
		private static void DeleteFile(FileActionItem item)
		{
			int count = 0;
			string filename = "";
			List<string> filenames = new List<string>();
			int index = 0;
			char[] wildcards = new char[] { '*', '?' };
			string workingPath = "";

			if(item != null)
			{
				filename = GetPropertyByName(item, "Filename");
				if(filename.Length > 0)
				{
					Console.WriteLine($" {filename}");
					workingPath = GetPropertyByName(item, nameof(WorkingPath));
					filenames = ResolveWildcards(workingPath, filename);
					count = filenames.Count;
					foreach(string filenameItem in filenames)
					{
						try
						{
							File.Delete(filenameItem);
							index++;
						}
						catch(Exception ex)
						{
							Console.WriteLine(" Error: Couldn't delete " +
								$"{Path.GetFileName(filenameItem)} - {ex.Message}");
						}
					}
					Console.WriteLine($" {index} of {count} files deleted...");
				}
				else
				{
					Console.WriteLine(" Filename property was not specified...");
					item.Stop = true;
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DelEveryX																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Delete every Xth file in the input files list.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void DelEveryX(FileActionItem item)
		{
			int count = 0;
			List<FileInfo> filesMatch = null;
			float fIndex = 0f;

			if(item != null)
			{
				//	Delete every Xth file in the input files list.
				if(CheckElements(item,
					ActionElementEnum.Inputs |
					ActionElementEnum.Count))
				{
					//	After the files have been tested for input, all of the
					//	wildcards will have been resolved.
					filesMatch = new List<FileInfo>();
					fIndex = 0f;
					foreach(FileInfo fileInfoItem in item.InputFiles)
					{
						fIndex++;
						if(fIndex >= item.Count)
						{
							filesMatch.Add(fileInfoItem);
							fIndex -= item.Count;
						}
					}
					//InputFiles.Clear();
					IdentifyInputFiles(item);
					count = filesMatch.Count;
					while(filesMatch.Count > 0)
					{
						Console.WriteLine($" {filesMatch[0].Name}");
						filesMatch[0].Delete();
						filesMatch.RemoveAt(0);
					}
					Console.WriteLine($" {count} files deleted...");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DirToTsv																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run a directory to a text-separated values (TSV) file.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void DirToTsv(FileActionItem item)
		{
			bool bLineInit = false;
			bool bRecurse = true;
			bool bSeparate = false;
			StringBuilder content = new StringBuilder();
			List<string> fieldList = new List<string>();
			List<FileInfo> fileList = new List<FileInfo>();
			StringBuilder line = new StringBuilder();
			string[] names = null;
			FileOptionItem optFields = null;
			FileOptionItem optRecurse = null;
			FileOptionItem optSepFileExt = null;

			if(item != null)
			{
				//	Take a directory reading of the input folder and output
				//	the information to a TSV file.
				if(CheckElements(item,
					ActionElementEnum.InputFolderName |
					ActionElementEnum.OutputFilename))
				{
					//	After the files have been tested for input, all of the
					//	wildcards will have been resolved.
					//	Options here are:
					//	Fields,{Filename;Extension;
					//	 RelativeDir;AbsoluteDir,DateTime}
					//	  Default Filename;RelativeDir
					//	Recurse,{true|false} - Default true
					//	SeparateFilenameExtension,{true|false} - Default false
					optFields = GetOptionByName(item, "Fields");
					optRecurse = GetOptionByName(item, "Recurse");
					optSepFileExt = GetOptionByName(item, "SeparateFilenameExtension");

					if(optSepFileExt != null)
					{
						bSeparate = ToBool(optSepFileExt.Value);
					}

					if(optFields != null)
					{
						names = optFields.Value.Split(';',
							StringSplitOptions.RemoveEmptyEntries |
							StringSplitOptions.TrimEntries);
					}
					if(names?.Length > 0)
					{
						foreach(string fieldNameItem in names)
						{
							fieldList.Add(fieldNameItem);
						}
					}
					else
					{
						fieldList.Add("Filename");
						if(bSeparate)
						{
							fieldList.Add("Extension");
						}
						fieldList.Add("RelativeDir");
					}

					if(optRecurse != null)
					{
						bRecurse = ToBool(optRecurse.Value);
					}

					FillFileList(fileList, item.InputDir, bRecurse);
					foreach(FileInfo fileItem in fileList)
					{
						//	TODO: Allow fields to be output in any order.
						bLineInit = false;
						Clear(line);
						//	DateTime.
						if(fieldList.Exists(x => x.ToLower() == "datetime"))
						{
							if(bLineInit)
							{
								line.Append('\t');
							}
							line.Append(
								fileItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"));
							bLineInit = true;
						}
						//	Filename.
						if(fieldList.Exists(x => x.ToLower() == "filename"))
						{
							if(bLineInit)
							{
								line.Append('\t');
							}
							if(bSeparate)
							{
								line.Append(Path.GetFileNameWithoutExtension(fileItem.Name));
							}
							else
							{
								line.Append(fileItem.Name);
							}
							bLineInit = true;
						}
						//	Extension.
						if(fieldList.Exists(x => x.ToLower() == "extension"))
						{
							if(bLineInit)
							{
								line.Append('\t');
							}
							line.Append(Path.GetExtension(fileItem.Name).Substring(1));
							bLineInit = true;
						}
						//	Relative Directory.
						if(fieldList.Exists(x => x.ToLower() == "relativedir"))
						{
							if(bLineInit)
							{
								line.Append('\t');
							}
							line.Append(
								GetRelativeDirectory(
									item.InputDir.FullName,
									fileItem.Directory.FullName));
							bLineInit = true;
						}
						//	Absolute Directory.
						if(fieldList.Exists(x => x.ToLower() == "absolutedir"))
						{
							if(bLineInit)
							{
								line.Append('\t');
							}
							line.Append(fileItem.Directory.FullName);
							bLineInit = true;
						}
						content.AppendLine(line.ToString());
					}
					File.WriteAllText(item.OutputFile.FullName, content.ToString());
					Console.WriteLine($" {fileList.Count} files listed...");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* DrawImage																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Draw the image specified by ImageName onto the working image at the
		/// location specified by user properties Left and Top.
		/// </summary>
		/// <param name="item">
		/// Reference to the action item describing the image to draw and the
		/// location at which to draw it.
		/// </param>
		private static void DrawImage(FileActionItem item)
		{
			Bitmap bitmap = null;
			int height = 0;
			BitmapInfoItem sourceImage = null;
			Rectangle sourceRect = Rectangle.Empty;
			Rectangle targetRect = Rectangle.Empty;
			int width = 0;
			int x = 0;
			int y = 0;

			if(item != null && WorkingImage != null)
			{
				sourceImage = Images.FirstOrDefault(x =>
					x.Name == GetPropertyByName(item, "ImageName"));
				if(sourceImage != null)
				{
					Console.WriteLine($" {sourceImage.Name}");
					bitmap = sourceImage.Bitmap;
					sourceRect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
					x = ToInt(GetPropertyByName(item, "Left"));
					y = ToInt(GetPropertyByName(item, "Top"));
					width = ToInt(GetPropertyByName(item, "Width"));
					height = ToInt(GetPropertyByName(item, "Height"));
					if(width == 0)
					{
						width = bitmap.Width;
					}
					if(height == 0)
					{
						height = bitmap.Height;
					}
					targetRect = new Rectangle(x, y, width, height);
					using(Graphics graphics = Graphics.FromImage(WorkingImage.Bitmap))
					{
						InitializeGraphics(graphics);
						graphics.DrawImage(bitmap,
							targetRect, sourceRect, GraphicsUnit.Pixel);
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FileOpenImage																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Open the image specified in InputFilename.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which the item will be opened.
		/// </param>
		/// <remarks>
		/// This method works upon the currently open file.
		/// </remarks>
		private static void FileOpenImage(FileActionItem item)
		{
			Bitmap bitmap = null;
			Bitmap bitmapA = null;
			FileInfo file = null;
			string imageFilename = "";
			BitmapInfoItem imageInfo = null;
			string imageName = "";

			if(item != null)
			{
				imageFilename = GetPropertyByName(item, "ImageFilename");
				if(imageFilename.Length > 0)
				{
					file = new FileInfo(AbsolutePath(item.WorkingPath, imageFilename));
				}
				else
				{
					file = GetCurrentFile(item);
				}
				if(file != null && file.Exists)
				{
					imageName = GetPropertyByName(item, "ImageName");
					if(imageName.Length == 0)
					{
						//imageName = $"Image{PadLeft("0", Images.Count, 5)}";
						imageName = file.Name;
					}
					Console.WriteLine($" {imageName}");
					bitmap = (Bitmap)Bitmap.FromFile(file.FullName);
					bitmapA = new Bitmap(bitmap.Width, bitmap.Height,
						System.Drawing.Imaging.PixelFormat.Format32bppArgb);
					using(Graphics gr = Graphics.FromImage(bitmapA))
					{
						gr.DrawImage(bitmap,
							new Rectangle(0, 0, bitmapA.Width, bitmapA.Height));
					}
					bitmap.Dispose();
					bitmap = bitmapA;
					imageInfo = Images.FirstOrDefault(x => x.Name == imageName);
					if(imageInfo == null)
					{
						imageInfo = new BitmapInfoItem()
						{
							Name = imageName
						};
						Images.Add(imageInfo);
					}
					imageInfo.Bitmap = bitmap;
					if(ToBool(GetPropertyByName(item, "IsWorkingImage"), true))
					{
						WorkingImage = imageInfo;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FileOverlayImage																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Open each image from the range and place the image specified in
		/// InputFilename at the options specified by Left, Top, Width, and Height.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which the item will be opened.
		/// </param>
		private static void FileOverlayImage(FileActionItem item)
		{
			bool bContinue = true;
			bool bRange = false;
			byte[] bytes = null;
			Graphics g = null;
			int height = 0;
			int left = 0;
			Bitmap maskBitmap = null;
			FileInfo maskFile = null;
			List<string> names = null;
			FileOptionItem optionHeight = null;
			FileOptionItem optionLeft = null;
			FileOptionItem optionMask = null;
			FileOptionItem optionTop = null;
			FileOptionItem optionWidth = null;
			Bitmap sourceBitmap = null;
			FileInfo sourceFile = null;
			List<string> sourceFilenames = new List<string>();
			int top = 0;
			int width = 0;

			if(item != null)
			{
				optionMask = GetOptionByName(item, "MaskFilename");
				if(optionMask != null)
				{
					maskFile = new FileInfo(optionMask.Value);
					if(!maskFile.Exists)
					{
						maskFile = null;
					}
				}
				optionLeft = GetOptionByName(item, "Left");
				optionTop = GetOptionByName(item, "Top");
				optionWidth = GetOptionByName(item, "Width");
				optionHeight = GetOptionByName(item, "Height");
				if(item.InputFiles.Count > 0)
				{
					bContinue = true;
				}
				else
				{
					bRange = bContinue = CheckElements(item,
						ActionElementEnum.InputFolderName |
						ActionElementEnum.Range);
				}
				if(bContinue &&
					maskFile != null &&
					optionLeft != null && optionTop != null &&
					optionWidth != null && optionHeight != null)
				{
					maskBitmap = (Bitmap)Bitmap.FromFile(maskFile.FullName);
					left = ToInt(optionLeft.Value);
					top = ToInt(optionTop.Value);
					width = ToInt(optionWidth.Value);
					height = ToInt(optionHeight.Value);
					if(bRange)
					{
						//	Range-based.
						names = EnumerateRange(item.Range, item.Digits);
						foreach(string nameItem in names)
						{
							sourceFilenames.Add(
								Path.Combine(item.InputDir.FullName, nameItem));
						}
					}
					else
					{
						sourceFilenames = new List<string>();
						foreach(FileInfo fileInfoItem in item.InputFiles)
						{
							sourceFilenames.Add(fileInfoItem.FullName);
						}
					}
					if(sourceFilenames.Count > 0)
					{
						//	Source filenames were generated.
						foreach(string sourceFilenameItem in sourceFilenames)
						{
							sourceFile = new FileInfo(sourceFilenameItem);
							if(sourceFile.Exists)
							{
								bytes = File.ReadAllBytes(sourceFile.FullName);
								using(var ms = new MemoryStream(bytes))
								{
									sourceBitmap = new Bitmap(ms);
								}
								g = Graphics.FromImage(sourceBitmap);
								g.CompositingQuality =
									System.Drawing.Drawing2D.CompositingQuality.HighQuality;
								g.InterpolationMode =
									System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
								g.PixelOffsetMode =
									System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
								g.SmoothingMode =
									System.Drawing.Drawing2D.SmoothingMode.HighQuality;
								g.DrawImage(maskBitmap, left, top, width, height);
								g.Dispose();
								//sourceFile.Delete();
								sourceBitmap.Save(sourceFile.FullName);
								Console.WriteLine($" {Path.GetFileName(sourceFilenameItem)}");
							}
						}
					}
					else
					{
						Console.Write(" Error: Source filenames could not be ");
						Console.WriteLine("enumerated from the given range.");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FileSaveImage																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Save the working image to the specified OutputFilename.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which the item will be opened.
		/// </param>
		private static void FileSaveImage(FileActionItem item)
		{
			FileInfo file = null;

			if(item != null && WorkingImage != null)
			{
				if(item.OutputFile != null)
				{
					file = item.OutputFile;
				}
				else if(item.CurrentFile != null && item.OutputDir != null)
				{
					file = new FileInfo(
						Path.Combine(item.OutputDir.FullName, item.CurrentFile.Name));
				}
				if(file != null)
				{
					AssureFolder(file.Directory.FullName, true, quiet: true);
					WorkingImage.Bitmap.Save(file.FullName);
					Console.WriteLine($" File saved: {file.Name}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FillFileList																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Fill a caller-supplied list of files in the specified directory.
		/// </summary>
		/// <param name="fileList">
		/// Reference to the list of files to fill.
		/// </param>
		/// <param name="dir">
		/// Reference to the directory to check.
		/// </param>
		/// <param name="recurse">
		/// Value indicating whether to recurse to child folders.
		/// </param>
		/// <returns>
		/// Reference to a list of files matching the caller's specification.
		/// </returns>
		private static void FillFileList(List<FileInfo> fileList,
			DirectoryInfo dir, bool recurse)
		{
			DirectoryInfo[] dirs = null;
			FileInfo[] files = null;

			if(fileList != null && dir != null)
			{
				files = dir.GetFiles();
				if(files?.Length > 0)
				{
					fileList.AddRange(files);
				}
				if(recurse)
				{
					dirs = dir.GetDirectories();
					foreach(DirectoryInfo dirItem in dirs)
					{
						FillFileList(fileList, dirItem, recurse);
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FindFiles																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Find files with regular expression find pattern.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void FindFiles(FileActionItem item)
		{
			string find = "";
			List<FileInfo> filesMatch = null;

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFolderName))
				{
					if(item.Properties.Exists(x => x.Name == "Find"))
					{
						find = item.Properties.First(x => x.Name == "Find").Value;
					}
					if(find.Length > 0)
					{
						filesMatch = new List<FileInfo>();
						FindFilesRecursive(item.InputDir, find, item.Recurse);
					}
					else
					{
						if(find.Length == 0)
						{
							Console.WriteLine(" Error: Find property was not specified.");
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FindFilesRecursive																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Performs a recursive file search operation based upon all of the
		/// command parameters that have already been set.
		/// </summary>
		/// <param name="dir">
		/// Reference to the directory within which the files will be found.
		/// </param>
		/// <param name="find">
		/// Regular expression pattern to find.
		/// </param>
		/// <param name="recurse">
		/// Value indicating whether to recurse into sub-folders.
		/// </param>
		private static void FindFilesRecursive(DirectoryInfo dir,
			string find, bool recurse)
		{
			bool bDirectoryWritten = false;
			DirectoryInfo[] dirs = null;
			List<FileInfo> files = null;

			if(dir != null && find?.Length > 0)
			{
				if(dir.Name.ToLower().IndexOf("svg") > -1)
				{
					Trace.WriteLine("Break here: FileTools.FindFilesRecursive...");
				}
				if(Regex.IsMatch(dir.Name, find))
				{
					Console.WriteLine($"{dir.FullName}");
					bDirectoryWritten = true;
				}
				files = dir.GetFiles().ToList();
				foreach(FileInfo fileInfo in files)
				{
					if(Regex.IsMatch(fileInfo.Name, find))
					{
						if(!bDirectoryWritten)
						{
							Console.WriteLine($"{dir.FullName}");
							bDirectoryWritten = true;
						}
						Console.Write($" {fileInfo.Name.PadRight(50, '.')} ");
						Console.Write($"{fileInfo.LastWriteTime:yyyy-MM-dd} ");
						Console.WriteLine(
							$"{fileInfo.Length.ToString().PadLeft(10, ' ')}");
					}
				}
				if(bDirectoryWritten)
				{
					Console.WriteLine();
				}
				if(recurse)
				{
					dirs = dir.GetDirectories();
					if(dirs.Length > 0)
					{
						foreach(DirectoryInfo dirItem in dirs)
						{
							FindFilesRecursive(dirItem, find, recurse);
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ForEachFile																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the Actions collection of the presented object through all of the
		/// files in this item's InputFiles collection using the CurrentFile
		/// property for each one.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item representing the loop base.
		/// </param>
		private static async void ForEachFile(FileActionItem item)
		{
			if(item != null)
			{
				foreach(FileInfo fileItem in item.InputFiles)
				{
					item.CurrentFile = fileItem;
					await RunActions(item.Actions);
					//foreach(FileActionItem actionItem in item.Actions)
					//{
					//	await actionItem.Run();
					//}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FormatDirFile																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Reformat the contents of a DIR /s file, outputting it to a new CSV
		/// file.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item representing the loop base.
		/// </param>
		private static void FormatDirFile(FileActionItem item)
		{
			bool bDir = false;
			StringBuilder builder = new StringBuilder();
			string cDir = "";
			string content = "";
			string filename = "";
			string fDate = "";
			string fName = "";
			string fSize = "";
			string line = "";
			string pathName = "";
			string text = "";

			Console.WriteLine($" Item {(item == null ? "in" : "")}valid.");
			if(item != null)
			{
				Console.WriteLine($" Input file count: {item.InputFiles.Count}");
				Console.WriteLine($" Output filename: {item.OutputFilename}");
			}
			if(item != null &&
				item.InputFiles.Count > 0 &&
				item.OutputFilename?.Length > 0)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFilename | ActionElementEnum.OutputFilename))
				{
					//	Input and output filenames were both provided.
					filename = item.InputFiles[0].FullName;
					content = File.ReadAllText(filename);
					using(TextReader reader = new StringReader(content))
					{
						line = reader.ReadLine();
						while(line != null)
						{
							if(line.Length > 14)
							{
								if(line.StartsWith(" Directory of "))
								{
									//	Get the current path.
									pathName = Right(line, line.Length - 14);
								}
								else if(IsNumeric(line.Substring(0, 2)) &&
									line.Substring(2, 1) == "/" &&
									Right(line, 2) != " ." &&
									Right(line, 2) != "..")
								{
									if(line.Length > 36)
									{
										//	Current file or directory name.
										//01/02/2020  06:18    <DIR>          .
										//01/02/2020  06:18    <DIR>          ..
										//06/18/2018  08:07    <DIR>          Accounting
										//07/24/2014  17:07                72 Angles.xls
										//11/15/2019  11:12           241,472 Angles.xlsm
										bDir = line.Substring(21, 5) == "<DIR>";
										if(bDir)
										{
											cDir = "D";
										}
										else
										{
											cDir = "F";
										}
										fName = Right(line, line.Length - 36);
										if(bDir)
										{
											fSize = "";
										}
										else
										{
											text = line.Substring(19, 16);
											text = text.Replace(" ", "");
											text = text.Replace(",", "");
											fSize = text;
										}
										fDate = line.Substring(6, 4) +
											line.Substring(0, 2) +
											line.Substring(3, 2) + "." +
											line.Substring(12, 2) +
											line.Substring(15, 2);
										builder.Append(
											$"{fDate}\t{cDir}\t{fSize}\t{pathName}\t");
										builder.AppendLine(fName);
									}
								}
							}
							line = reader.ReadLine();
						}
					}
					File.WriteAllText(item.OutputFilename, builder.ToString());
					Console.WriteLine($" File created: {item.OutputFile.Name}");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetCurrentFile																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the current file.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action for which the current file will be
		/// retrieved.
		/// </param>
		/// <returns>
		/// Reference to the current file in focus, if found. Otherwise, null.
		/// </returns>
		/// <remarks>
		/// If a file has been placed in focus by a loop or other method, that
		/// object will be returned from the CurrentFile property. Otherwise, the
		/// first item in InputFiles collection is returned.
		/// </remarks>
		private static FileInfo GetCurrentFile(FileActionItem item)
		{
			FileInfo result = null;

			if(item != null)
			{
				//	An item has been provided.
				//	Test order.
				//	-	Local current file.
				//	- Local first item.
				//	- Parent current file.
				//	- Parent first item.
				if(item.mCurrentFile != null)
				{
					result = item.mCurrentFile;
				}
				else if(item.mInputFiles.Count > 0)
				{
					result = item.mInputFiles[0];
				}
				else
				{
					result = item.CurrentFile;
					if(result == null && item.InputFiles.Count > 0)
					{
						result = item.InputFiles[0];
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetRoot																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the root parent of the action specified.
		/// </summary>
		/// <param name="item">
		/// Reference to the item for which the root parent will be found.
		/// </param>
		/// <returns>
		/// Reference to the root parent, if found. Otherwise, null.
		/// </returns>
		private static FileActionItem GetRoot(FileActionItem item)
		{
			FileActionItem result = null;

			if(item != null)
			{
				if(item.mParent != null && item.mParent.Parent != null)
				{
					result = GetRoot(item.mParent.Parent);
				}
				else
				{
					result = item;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IdentifyInputFiles																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Identify the input and output files and directories at the current
		/// level for the specified item.
		/// </summary>
		/// <param name="item">
		/// Reference to the item being fulfilled.
		/// </param>
		/// <remarks>
		/// <para>
		/// When this method is called, make sure that the InitializeLevels
		/// method has already been called.
		/// </para>
		/// <para>
		/// Only call CheckElements after first calling IdentifyFiles. That
		/// method relies on the file objects in this version.
		/// </para>
		/// <para>
		/// In this version, only the local InputFiles group is resolved. If you
		/// want to use a globally resolvable template, implement a user property
		/// containing that template name and set the InputFilename, etc., property
		/// at the site with a reference to that custom property.
		/// </para>
		/// </remarks>
		private static void IdentifyInputFiles(FileActionItem item)
		{
			DirectoryInfo dir = null;
			FileInfo file = null;
			string filename = "";
			bool result = true;

			if(item != null && item.mInputNames.Count > 0)
			{
				//	Working path.
				if(item.WorkingPath.Length > 0)
				{
					dir = new DirectoryInfo(
						GetPropertyByName(item, nameof(WorkingPath)));
					if(!dir.Exists)
					{
						Console.WriteLine(" Error: Working path does not exist.");
						result = false;
					}
					else if((dir.Attributes & FileAttributes.Directory) !=
						FileAttributes.Directory)
					{
						Console.Write(" Error: A file was specified as the working ");
						Console.WriteLine("directory.");
						result = false;
					}
				}

				//	Input.
				item.mInputDir = null;
				item.mInputFiles.Clear();

				if(result)
				{
					if(item.mInputNames.Count > 0)
					{
						//	Input files are present.
						foreach(string filenameItem in item.mInputNames)
						{
							filename = AbsolutePath(
								GetPropertyByName(item, nameof(WorkingPath)),
								NormalizeValue(item, filenameItem));
							if(filename.Length > 0)
							{
								//	A filename has been retrieved.
								//	Check for wildcards and resolve variables.
								item.mInputFiles.AddRange(
									ResolveFilename(filename, false));
							}
						}
						if(item.mInputFiles.Count > 0)
						{
							file = item.mInputFiles[0];
							if((file.Attributes & FileAttributes.Directory) ==
								(FileAttributes)0)
							{
								//	This item is a file.
								item.mInputDir = new DirectoryInfo(file.Directory.FullName);
							}
							else
							{
								//	This item is a directory.
								item.mInputDir = new DirectoryInfo(file.FullName);
							}
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IdentifyOutputFiles																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Identify the output files and directories at the current
		/// level for the specified item.
		/// </summary>
		/// <param name="item">
		/// Reference to the item being fulfilled.
		/// </param>
		/// <remarks>
		/// <para>
		/// When this method is called, make sure that the InitializeLevels
		/// method has already been called.
		/// </para>
		/// <para>
		/// Only call CheckElements after first calling IdentifyFiles. That
		/// method relies on the file objects in this version.
		/// </para>
		/// </remarks>
		private static void IdentifyOutputFiles(FileActionItem item)
		{
			DirectoryInfo dir = null;
			FileInfo file = null;
			string filename = "";
			List<FileInfo> files = null;
			bool result = true;

			if(item != null)
			{
				//	Working path.
				if(item.WorkingPath.Length > 0)
				{
					dir = new DirectoryInfo(
						GetPropertyByName(item, nameof(WorkingPath)));
					if(!dir.Exists)
					{
						Console.WriteLine(" Error: Working path does not exist.");
						result = false;
					}
					else if((dir.Attributes & FileAttributes.Directory) !=
						FileAttributes.Directory)
					{
						Console.Write(" Error: A file was specified as the working ");
						Console.WriteLine("directory.");
						result = false;
					}
				}

				//	Output.
				item.OutputDir = null;
				item.OutputFile = null;

				if(result)
				{
					if(item.OutputName?.Length > 0 && item.IsOutputLocal())
					{
						//	Output folder or file is present.
						files = new List<FileInfo>();
						filename = AbsolutePath(
							GetPropertyByName(item, nameof(WorkingPath)),
							GetPropertyByName(item, nameof(OutputName)));
						files.AddRange(ResolveFilename(filename, true));
						if(files.Count > 0)
						{
							file = files[0];
							if((!file.Exists && file.Extension.Length > 0) ||
								(file.Exists &&
								((file.Attributes & FileAttributes.Directory) ==
								(FileAttributes)0)))
							{
								//	This item is a file.
								item.OutputFile = file;
								item.OutputDir = new DirectoryInfo(file.Directory.FullName);
							}
							else
							{
								//	This item is a directory.
								item.OutputDir = new DirectoryInfo(file.FullName);
							}
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* If																																		*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run one or more sets of actions if their conditions are true.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item for which this action is being
		/// called.
		/// </param>
		private static async void If(FileActionItem item)
		{
			bool bMatch = false;
			ConditionCollection conditions = null;
			bool conditionResult = false;
			ExpressionContext context;
			IDynamicExpression dynCondition = null;

			if(item != null)
			{
				context = new ExpressionContext();
				//// Allow the expression to use all static public methods of
				//// System.Math.
				//context.Imports.AddType(typeof(Math));
				context.Variables["CurrentFilename"] = item.CurrentFile.Name;
				context.Variables["CurrentFileNumber"] =
					GetIndexValue(item.CurrentFile.Name);

				foreach(FileActionItem actionItem in item.Actions)
				{
					if(!actionItem.Options.Exists(x => x.Name.ToLower() == "mute"))
					{
						conditions = GetConditions(actionItem);
						bMatch = true;
						foreach(ConditionItem conditionItem in conditions)
						{
							dynCondition = context.CompileDynamic(conditionItem.Condition);
							conditionResult = (bool)dynCondition.Evaluate();
							if(!conditionResult)
							{
								bMatch = false;
								break;
							}
						}
						if(bMatch)
						{
							//	This item evaluates to true. Run its actions.
							await RunActions(actionItem.Actions);
							//foreach(FileActionItem trueActionItem in actionItem.Actions)
							//{
							//	await trueActionItem.Run();
							//}
						}
					}
					else
					{
						Console.WriteLine($"Action {actionItem.Action} is muted...");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ImageBackground																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Set the background color or image on the working image.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item for which this action is being
		/// called.
		/// </param>
		private static void ImageBackground(FileActionItem item)
		{
			System.Drawing.Brush brush = null;
			System.Drawing.Color color = System.Drawing.Color.Empty;
			string backgroundColor = "";
			Bitmap backgroundBitmap = null;
			string backgroundFilename = "";
			Bitmap bitmap = null;
			Rectangle rect = Rectangle.Empty;
			int height = 0;
			int width = 0;

			if(item != null && WorkingImage?.Bitmap != null)
			{
				//	An item was presented and a working image is present.
				width = WorkingImage.Bitmap.Width;
				height = WorkingImage.Bitmap.Height;
				rect = new Rectangle(0, 0, width, height);
				backgroundBitmap = new Bitmap(width, height);
				using(Graphics graphics = Graphics.FromImage(backgroundBitmap))
				{
					//	Set the color first.
					InitializeGraphics(graphics);
					backgroundColor = GetPropertyByName(item, "BackgroundColor");
					if(backgroundColor.Length > 0)
					{
						brush = new SolidBrush(ColorTranslator.FromHtml(backgroundColor));
						graphics.FillRectangle(brush, rect);
					}
					//	Check for image.
					backgroundFilename = GetPropertyByName(item, "BackgroundImage");
					if(backgroundFilename.Length > 0)
					{
						bitmap = (Bitmap)Bitmap.FromFile(
							AbsolutePath(item.WorkingPath, backgroundFilename));
						graphics.DrawImage(bitmap,
							rect,
							new Rectangle(0, 0, bitmap.Width, bitmap.Height),
							GraphicsUnit.Pixel);
					}
					//	Paint the working image over the background.
					graphics.DrawImage(WorkingImage.Bitmap,
						rect, rect, GraphicsUnit.Pixel);
				}
				WorkingImage.Bitmap = backgroundBitmap;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ImagesClear																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Clear the contents of the Images collection.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item for which this action is being
		/// called.
		/// </param>
		/// <remarks>
		/// This method also clears the WorkingImage property.
		/// </remarks>
		private static void ImagesClear(FileActionItem item)
		{
			if(item != null)
			{
				Images.Clear();
				WorkingImage = null;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ImageSetCommonBoundary																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Create common bounding boxes for the image sets specified.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void ImageSetCommonBoundary(FileActionItem item)
		{
			string filename = "";
			FileOptionItem fileOption = null;
			Bitmap image = null;
			RectInfoItem rect = null;

			if(item != null)
			{
				//	Create common bounding boxes for the image sets specified.
				Console.WriteLine($" Text: {item.Text}");
				if(CheckElements(item,
					ActionElementEnum.Inputs))
				{
					//	After the files have been tested for input, all of the
					//	wildcards will have been resolved.
					image = ReadAllImages(item.InputFiles);
					rect = FindBoundingBox(image);
					fileOption = GetOptionByName(item, "Square");
					if(fileOption != null)
					{
						Console.WriteLine(" Squaring Image...");
						RectInfoItem.Square(rect);
					}
					fileOption = GetOptionByName(item, "Grow");
					if(fileOption != null && ToFloat(fileOption.Value) != 0f)
					{
						Console.WriteLine(
							$" Growing image by {ToFloat(fileOption.Value)}");
						RectInfoItem.Grow(rect, ToFloat(fileOption.Value));
					}
					fileOption = GetOptionByName(item, "SavePicture");
					if(fileOption != null && fileOption.Value?.Length > 0)
					{
						//	Save a copy of the processed image.
						filename = AbsolutePath(
							GetPropertyByName(item, nameof(WorkingPath)),
							NormalizeValue(item, fileOption.Value));
						image.Save(filename);
						Console.WriteLine($" Working image saved: {filename}.");
					}

					if(item.Text.Length > 0)
					{
						rect.Name = GetPropertyByName(item, nameof(Text));
					}
					RectangleInfoList.Add(rect);

					Console.WriteLine(rect);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* InitializeFilenames																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Initialize values for working at this and child levels.
		/// </summary>
		/// <param name="item">
		/// Reference to the item to be intialized.
		/// </param>
		/// <remarks>
		///	<para>When preparing the object for use:</para>
		///	<list type="bullet">
		///	<item>All input files at a level should be read from a single
		///	reference source. Check for an inputs collection.</item>
		///	<item>If blank, check for filename and add to inputs
		///	collection.</item>
		///	<item>If blank, check for foldername and add to inputs
		///	collection.</item>
		///	<item>All output files at a level should be written from a
		///	single reference source. Check for the Output.</item>
		///	<item>If blank, check for the output filename.</item>
		///	<item>If blank, check for the output foldername.</item>
		///	</list>
		///	<para>
		///	In this version, the conversion is made on every action at every level.
		///	</para>
		/// </remarks>
		private static void InitializeFilenames(FileActionItem item)
		{
			if(item != null)
			{
				//	Input filenames.
				if(item.mInputNames.Count == 0 &&
					(item.mInputFilename?.Length > 0 ||
					item.mInputFolderName?.Length > 0))
				{
					//	The input names collection was not specified, but either a
					//	filename or foldername were provided.
					if(item.mInputFilename?.Length > 0)
					{
						//	An input filename was provided at this level.
						item.mInputNames.AddRange(
							ResolveWildcards(
								GetPropertyByName(item,
									nameof(WorkingPath)), item.mInputFilename));
						//item.mInputNames.Add(item.mInputFilename);
						item.mInputFilename = "";
					}
					if(item.mInputFolderName?.Length > 0)
					{
						//	An input foldername was provided at this level.
						item.mInputNames.Add(item.mInputFolderName);
						item.mInputFiles.Add(new FileInfo(item.mInputFolderName));
						//	DEP20240225.1102 - I don't remember the original reason
						//	for clearing this variable. Its raw value is needed in
						//	directory deletion routine.
						//	Be aware this may need to be uncommented.
						//item.mInputFolderName = "";
					}
				}
				//	Output filenames.
				if((item.mOutputName == null || item.mOutputName.Length == 0) &&
					(item.mOutputFilename?.Length > 0 ||
					item.mOutputFolderName?.Length > 0))
				{
					//	The output name was not specified and either an output filename
					//	or output foldername are present.
					if(item.mOutputFilename?.Length > 0)
					{
						//	An output filename was provided at this level.
						item.mOutputName = item.mOutputFilename;
					}
					else if(item.mOutputFolderName?.Length > 0)
					{
						//	An output folder name was provided at this level.
						item.mOutputName = item.mOutputFolderName;
					}
				}
				//	In this version, all child items are processed as they are
				//	encountered.
				////	Process all child levels.
				//foreach(FileActionItem actionItem in item.mActions)
				//{
				//	InitializeFilenames(actionItem);
				//}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* InitializeProperties																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Initialize the public properties list of this class so they can be
		/// used repeatedly with minimal overhead.
		/// </summary>
		private static void InitializeProperties()
		{
			BindingFlags bindingFlags = BindingFlags.Instance | BindingFlags.Public;
			PropertyInfo[] properties = null;

			if(mPublicProperties.Count == 0)
			{
				//	Only initialize once.
				properties = typeof(FileActionItem).GetProperties(bindingFlags);
				foreach(PropertyInfo propertyInfoItem in properties)
				{
					mPublicProperties.Add(propertyInfoItem);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* LoadRectangleInfoList																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Load the contents of the rectangle info list from an external JSON
		/// file.
		/// </summary>
		/// <param name="item">
		/// Reference to the active file action item from which the list is being
		/// loaded.
		/// </param>
		/// <remarks>
		/// This method uses the current input file.
		/// </remarks>
		private static void LoadRectangleInfoList(FileActionItem item)
		{
			string content = "";
			FileInfo file = null;
			RectInfoCollection rectangles = null;

			RectangleInfoList.Clear();
			if(item != null)
			{
				file = GetCurrentFile(item);
				if(file != null && file.Exists)
				{
					Console.WriteLine($" {file.Name}");
					content = File.ReadAllText(file.FullName);
					rectangles =
						JsonConvert.DeserializeObject<RectInfoCollection>(content);
					RectangleInfoList.AddRange(rectangles);
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* MaskRectangleBytes																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Paint a mask rectangle over the specified area of the working image.
		/// </summary>
		/// <param name="item">
		/// Reference to the active file action item from which the list is being
		/// loaded.
		/// </param>
		private static void MaskRectangleBytes(FileActionItem item)
		{
			//	TODO: This should probably be renamed to DrawRectangleBytes.
			Brush brush = null;
			string color = "";
			int height = 0;
			int width = 0;
			int x = 0;
			int y = 0;

			if(item != null && WorkingImage?.Bitmap != null)
			{
				color = GetPropertyByName(item, "Color");
				x = ToInt(GetPropertyByName(item, "Left"));
				y = ToInt(GetPropertyByName(item, "Top"));
				width = ToInt(GetPropertyByName(item, "Width"));
				height = ToInt(GetPropertyByName(item, "Height"));
				using(Graphics graphics = Graphics.FromImage(WorkingImage.Bitmap))
				{
					InitializeGraphics(graphics);
					brush = new SolidBrush(ColorTranslator.FromHtml(color));
					graphics.FillRectangle(brush,
						new Rectangle(x, y, width, height));
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* MoveFiles																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Move the selected files from one folder to another.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which variables will be resolved.
		/// </param>
		private static void MoveFiles(FileActionItem item)
		{

			if(item != null)
			{
				if(CheckElements(item, ActionElementEnum.Inputs))
				{
					//	After the files have been tested with the Input option, all of
					//	the wildcards will have been resolved.
					if(item.OutputFolderName.Length == 0)
					{
						//	If no output folder was specified, the target is the
						//	working folder.
						item.OutputFolderName =
							GetPropertyByName(item, nameof(WorkingPath));
					}
					if(item.OutputFolderName.Length > 0)
					{
						foreach(FileInfo fileInfoItem in item.InputFiles)
						{
							fileInfoItem.MoveTo(
								Path.Combine(item.OutputFolderName, fileInfoItem.Name));
							Console.WriteLine($" {fileInfoItem.Name} moved...");
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* NonLinearEditExcel																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Perform non-linear editing on the specified excel file.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which the variables will be resolved.
		/// </param>
		private static async Task NonLinearEditExcel(FileActionItem item)
		{
			FileInfo activeTarget = null;
			bool bFirst = false;
			Bitmap bitmap = null;
			BitmapInfoItem bitmapInfo = null;
			//Brush brush = null;
			int count = 0;
			IntRangeCollection cutRanges = null;
			DataSet data = null;
			List<NleEntryItem> edits = null;
			NleEntryCollection entries = null;
			List<NleEntryItem> entryItems = null;
			string ext = "";
			string filename = "";
			int index = 0;
			int indexInput = 0;
			int indexInputMax = 0;
			int indexOutput = 0;
			InputOutputFileItem ioFile = null;
			InputOutputFileCollection ioFiles =
				new InputOutputFileCollection();
			NameValueItem prop = null;
			string sheetName = "";
			FileInfo sourceFile = null;
			List<FileInfo> sourceFiles = null;
			DataTable table = null;
			FileActionItem workingItem = null;

			//	Execute a non-linear editing pattern using an Excel file with
			//	the fields Start, Action, End, Count, X, Y, Width, Height,
			//	and Color.
			//	The input file is expected to be an Excel file.
			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFilename |
					ActionElementEnum.SourceFolderName |
					ActionElementEnum.OutputFoldername))
				{
					//	In this case, the input filename is the Excel file,
					//	the input folder name is the location where the source files
					//	are found, and the output folder name is the location where the
					//	edited files will be stored.
					filename = item.InputFiles[0].FullName;
					data = ExcelFile.ReadWorkbook(filename, true);
					if(data != null && data.Tables.Count > 0)
					{
						prop = item.Properties.FirstOrDefault(x =>
							x.Name.ToLower() == "sheet");
						if(prop != null && data.Tables.Contains(prop.Value))
						{
							//	A sheet has been specified.
							sheetName = prop.Value;
						}
						else
						{
							sheetName = "EditSheet";
						}
						if(data.Tables.Contains(sheetName))
						{
							table = data.Tables[sheetName];
							sourceFiles = item.SourceDir.GetFiles().ToList();
							if(sourceFiles.Count > 0)
							{
								UpdateBaseAndDigits(item,
									sourceFiles[0].Name, sourceFiles.Count, 1);
								indexInput = 1;
								indexOutput = 1;
								entries = NleEntryCollection.FromDataTable(table);
								cutRanges = entries.GetCutRanges();
								//	Get the highest input filename index.
								foreach(FileInfo fileItem in sourceFiles)
								{
									indexInputMax =
										Math.Max(indexInputMax, GetIndexValue(fileItem.Name));
								}
								//	Remove all cut files from the source files list.
								Console.WriteLine(" Removing cut segments...");
								for(indexInput = 1; indexInput <= indexInputMax; indexInput ++)
								{
									if(cutRanges.Exists(x =>
										indexInput >= x.Start && indexInput <= x.End))
									{
										sourceFiles.RemoveAll(x =>
											x.Name ==
												FilePatternWithIndex(
													item.Base, indexInput, item.Digits));
									}
									if(indexInput % 1000 == 0)
									{
										Console.Write($", {indexInput}");
									}
								}
								//	Refresh the highest input filename index.
								indexInputMax = 0;
								foreach(FileInfo fileItem in sourceFiles)
								{
									indexInputMax =
										Math.Max(indexInputMax, GetIndexValue(fileItem.Name));
								}
								Console.WriteLine("");
								//	Create target files from input list.
								ioFiles = new InputOutputFileCollection(indexInputMax);
								Console.WriteLine(" Creating target files collection...");
								for(indexInput = 1; indexInput <= indexInputMax; indexInput ++)
								{
									if(indexInput % 1000 == 0)
									{
										Console.Write($", {indexInput}");
									}
									sourceFile = sourceFiles.FirstOrDefault(x =>
										x.Name == FilePatternWithIndex(item.Base,
											indexInput, item.Digits));
									if(sourceFile != null)
									{
										entryItems = entries.FindAll(x =>
											x.Action == NonLinearEditActionEnum.FreezeFrame &&
											x.Start == indexInput &&
											x.Count > 0);
										if(entryItems.Count > 0)
										{
											//	Freeze this frame for a specified number of places.
											bFirst = true;
											foreach(NleEntryItem entryItem in entryItems)
											{
												count = entryItem.Count;
												for(index = 0; index < count; index ++)
												{
													ioFiles.Add(new InputOutputFileItem()
													{
														FirstInSequence = bFirst,
														InputFile = sourceFile,
														SourceIndex = indexInput,
														TargetIndex = indexOutput,
														OutputFile = new FileInfo(
															AbsolutePath(item.WorkingPath,
																Path.Combine(item.OutputFolderName,
																	FilePatternWithIndex(item.Base,
																		indexOutput, item.Digits))))
													});
													indexOutput++;
													bFirst = false;
												}
											}
										}
										else
										{
											//	Place an individual mark.
											ioFiles.Add(new InputOutputFileItem()
											{
												InputFile = sourceFile,
												SourceIndex = indexInput,
												TargetIndex = indexOutput,
												OutputFile = new FileInfo(
													AbsolutePath(item.WorkingPath,
														Path.Combine(item.OutputFolderName,
															FilePatternWithIndex(item.Base,
																indexOutput, item.Digits))))
											});
											indexOutput++;
										}
									}
								}
								Console.WriteLine("");
								//	Apply all changes to target files.
								if(!item.OutputDir.Exists)
								{
									item.OutputDir.Create();
								}
								Console.WriteLine(" Applying updates...");
								count = ioFiles.Count;
								for(index = 0; index < count; index ++)
								{
									if(index > 0 && index % 1000 == 0)
									{
										Console.Write($", {index}");
									}
									//	TODO: Write first copy of each file, including changes.
									ioFile = ioFiles[index];
									if(ioFile.FirstInSequence)
									{
										//	This item is the first item in the sequence.
										Console.WriteLine("*****");
										Console.WriteLine($"File: {ioFile.InputFile.Name}");
										activeTarget = ioFile.OutputFile;
										edits = entries.FindAll(x =>
											x.Action != NonLinearEditActionEnum.Cut &&
											x.Action != NonLinearEditActionEnum.FreezeFrame &&
											ioFile.SourceIndex >= x.Start &&
											ioFile.SourceIndex <= NleEntryItem.GetEnd(x) &&
											!x.Mute);
										if(edits.Count > 0)
										{
											//	Edits need to be made for creating the target image.
											ext = Path.GetExtension(ioFile.InputFile.Name).ToLower();
											if(ext == ".bmp" ||
												ext == ".jpeg" ||
												ext == ".jpg" ||
												ext == ".png" ||
												ext == ".webp")
											{
												Images.Clear();
												bitmap = (Bitmap)Bitmap.FromFile(
													ioFile.InputFile.FullName);
												bitmapInfo = new BitmapInfoItem()
												{
													Bitmap = bitmap,
													Name = ioFile.InputFile.Name
												};
												Images.Add(bitmapInfo);
												WorkingImage = bitmapInfo;

												foreach(NleEntryItem editItem in edits)
												{
													workingItem = new FileActionItem();
													workingItem.WorkingPath = item.WorkingPath;
													workingItem.CurrentFile = ioFile.InputFile;
													if(editItem.Condition.Length > 0)
													{
														workingItem.Conditions.Add(new ConditionItem()
														{
															Condition = editItem.Condition,
															Assignment = editItem.Assignment
														});
													}
													workingItem.Properties.Add(new NameValueItem()
													{
														Name = "Color",
														Value = editItem.Color
													});
													workingItem.Properties.Add(new NameValueItem()
													{
														Name = "Left",
														Value = editItem.X.ToString()
													});
													workingItem.Properties.Add(new NameValueItem()
													{
														Name = "Top",
														Value = editItem.Y.ToString()
													});
													workingItem.Properties.Add(new NameValueItem()
													{
														Name = "Width",
														Value = editItem.Width.ToString()
													});
													workingItem.Properties.Add(new NameValueItem()
													{
														Name = "Height",
														Value = editItem.Height.ToString()
													});
													foreach(NameValueItem propertyItem in
														editItem.Properties)
													{
														workingItem.Properties.Add(propertyItem);
													}

													//	TODO: Fade in.
													//	TODO: Fade out.
													switch(editItem.Action)
													{
														case NonLinearEditActionEnum.AlphaConditionalAdjust:
															Console.WriteLine("Alpha conditional adjust");
															AlphaConditionalAdjustBytes(workingItem);
															break;
														case NonLinearEditActionEnum.AlphaMask:
															Console.WriteLine("Alpha mask");
															AlphaMask(workingItem);
															break;
														case NonLinearEditActionEnum.ColorConditionalAdjust:
															Console.WriteLine("Alpha conditional adjust");
															ColorConditionalAdjustBytes(workingItem);
															break;
														case NonLinearEditActionEnum.CropImage:
															Console.WriteLine("Crop image");
															CropImage(workingItem);
															break;
														case NonLinearEditActionEnum.ImageBackground:
															Console.WriteLine("Image background");
															ImageBackground(workingItem);
															break;
														case NonLinearEditActionEnum.MaskRectangle:
															if(editItem.Color.Length > 0 &&
																editItem.Width > 0 &&
																editItem.Height > 0)
															{
																Console.WriteLine("Mask rectangle");
																MaskRectangleBytes(workingItem);
															}
															else
															{
																Console.WriteLine(
																	$" Error on source frame {index}: Mask " +
																	"rectangle requires X, Y, Width, and " +
																	"Color\r\n");
															}
															break;
														case NonLinearEditActionEnum.RemoveBackground:
															Console.WriteLine("Remove background");
															await RemoveBackgroundBytes(workingItem);
															break;
														case NonLinearEditActionEnum.ResizeImage:
															Console.WriteLine("Resize image");
															ResizeImage(workingItem);
															break;
													}
												}
												bitmap = WorkingImage.Bitmap;
												bitmap.Save(ioFile.OutputFile.FullName);
												bitmap.Dispose();
												bitmap = null;
												WorkingImage = null;
											}
										}
										else
										{
											//	Straight file copy is available.
											ioFile.InputFile.CopyTo(
												ioFile.OutputFile.FullName, true);
										}
									}
									else if(activeTarget != null)
									{
										//	This item is subsequent. Copy the active target to
										//	this item's target.
										activeTarget.CopyTo(ioFile.OutputFile.FullName, true);
									}
								}
								Console.WriteLine("");
							}
						}
						else
						{
							Console.WriteLine($" Error: Sheet not found: {sheetName}");
						}
					}
					else
					{
						Console.WriteLine(
							$" Error: Could not read {Path.GetFileName(filename)}.");
						Console.WriteLine(" Check to see if you have it open elsewhere.");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* NormalizeValue																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Normalize the filename, using the values of any local properties
		/// necessary.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which variables will be resolved.
		/// </param>
		/// <param name="value">
		/// Value to normalize.
		/// </param>
		/// <returns>
		/// Fully normalized version of the provided filename.
		/// </returns>
		private static string NormalizeValue(FileActionItem item, string value)
		{
			MatchCollection matches = null;
			List<NameValueItem> replacements = new List<NameValueItem>();
			string result = "";

			if(value?.Length > 0)
			{
				result = value;
				matches = Regex.Matches(result, ResourceMain.rxEmbeddedFieldName);
				if(matches.Count > 0)
				{
					foreach(Match matchItem in matches)
					{
						if(!replacements.Exists(x =>
							x.Name == GetValue(matchItem, "field")))
						{
							replacements.Add(new NameValueItem()
							{
								Name = GetValue(matchItem, "field"),
								Value = GetPropertyByName(item, GetValue(matchItem, "name"))
							});
						}
					}
					foreach(NameValueItem replaceItem in replacements)
					{
						result = result.Replace(replaceItem.Name, replaceItem.Value);
					}
					//	Run at least one more time after having made replacements.
					result = NormalizeValue(item, result);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* PrefixFilenames																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Prefix the specified files.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void PrefixFilenames(FileActionItem item)
		{
			FileOptionItem fileOption = null;
			string name = "";
			string text = "";

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.Inputs |
					ActionElementEnum.OptionPrefix))
				{
					//	After the files have been tested for input, all of the
					//	wildcards will have been resolved.
					text = "";
					fileOption = GetOptionByName(item, "Prefix");
					if(fileOption != null)
					{
						text = fileOption.Value;
					}
					foreach(FileInfo fileInfoItem in item.InputFiles)
					{
						name = $"{text}{fileInfoItem.Name}";
						fileInfoItem.MoveTo(Path.Combine(
							fileInfoItem.Directory.FullName, name));
						Console.WriteLine($" {name}");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RemoveBackground																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Remove background from the selected files, outputting them to the
		/// specified target folder.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which variables will be resolved.
		/// </param>
		private static async Task RemoveBackground(FileActionItem item)
		{
			//Bitmap background = null;
			//Bitmap bitmap = null;
			StringBuilder builder = new StringBuilder();
			byte[] bytes = null;
			HttpClient client = null;
			//int count = 0;
			string endpoint = "";
			//Graphics graphics = null;
			Uri hostUri = null;
			string hostUrl = "";
			//int index = 0;
			string key = "";
			HttpResponseMessage response = null;

			if(item != null)
			{
				key = GetPropertyByName(item, "MSCognitiveServicesKey");
				endpoint = GetPropertyByName(item, "MSCognitiveServicesEndpoint");
				if(key.Length > 0 && endpoint.Length > 0)
				{
					//	The key and the endpoint were both provided.
					if(CheckElements(item, ActionElementEnum.Inputs))
					{
						//	After the files have been tested with the Input option, all of
						//	the wildcards will have been resolved.
						//if(item.OutputFolderName.Length == 0)
						//{
						//	//	If no output folder was specified, the target is the
						//	//	working folder.
						//	item.OutputFolderName =
						//		GetPropertyByName(item, nameof(WorkingPath));
						//}
						if(item.OutputFolderName.Length > 0)
						{
							if(!item.OutputDir.Exists)
							{
								item.OutputDir.Create();
							}
							client = new HttpClient();
							client.DefaultRequestHeaders.Add(
								"Ocp-Apim-Subscription-Key", key);
							//client.DefaultRequestHeaders.Remove("Content-Type");
							//client.DefaultRequestHeaders.Add(
							//	"Content-Type", "application/octet-stream");
							hostUrl = $"https://{endpoint}.cognitiveservices.azure.com" +
								"/computervision/imageanalysis:segment" +
								"?api-version=2023-02-01-preview" +
								"&mode=backgroundRemoval";
							hostUri = new Uri(hostUrl);
							foreach(FileInfo fileInfoItem in item.InputFiles)
							{
								bytes = File.ReadAllBytes(fileInfoItem.FullName);
								using(ByteArrayContent content = new ByteArrayContent(bytes))
								{
									//	Check to see if the default headers section works.
									content.Headers.Remove("Content-Type");
									content.Headers.Add(
										"Content-Type", "application/octet-stream");
									try
									{
										response = await client.PostAsync(hostUri, content);
										bytes = await response.Content.ReadAsByteArrayAsync();
										//using(MemoryStream stream = new MemoryStream(bytes))
										//{
										//	bitmap = new Bitmap(stream);
										//	bytes = ImageToColorBytes(bitmap);
										//	BlankGreenOverRatio(bytes, 255, 102, 128, 102);
										//	BlankGreenOverRatio(bytes, 255, 102, 143, 107);
										//	BlankGreenOverRatio(bytes, 255, 110, 140, 105);
										//	BlankGreenOverRatio(bytes, 255, 105, 143, 99);
										//	BlankGreenOverRatio(bytes, 255, 105, 135, 105);
										//	BlankAlphaLevel(bytes, 250);
										//	bitmap = ColorBytesToImage(bytes,
										//		bitmap.Width, bitmap.Height);
										//	//	Create the green background.
										//	bytes = new byte[bitmap.Width * bitmap.Height * 4];
										//	count = bytes.Length;
										//	for(index = 0; index < count; index += 4)
										//	{
										//		bytes[index + 1] = bytes[index + 3] = 0xff;
										//	}
										//	background = ColorBytesToImage(bytes,
										//		bitmap.Width, bitmap.Height);
										//	graphics = Graphics.FromImage(background);
										//	graphics.CompositingMode = CompositingMode.SourceOver;
										//	graphics.CompositingQuality =
										//		CompositingQuality.HighQuality;
										//	graphics.InterpolationMode =
										//		InterpolationMode.HighQualityBicubic;
										//	graphics.SmoothingMode = SmoothingMode.AntiAlias;
										//	graphics.DrawImage(bitmap, 0, 0);
										//	graphics.Dispose();
										//	bitmap = background;
										//	//	Store the image.
										//	bitmap.Save(
										//		Path.Combine(
										//			item.OutputDir.FullName, fileInfoItem.Name));
										//}
										// Use the the following save when not altering the movie.
										File.WriteAllBytes(
											Path.Combine(item.OutputDir.FullName, fileInfoItem.Name),
											bytes);
										Console.WriteLine(
											$" {fileInfoItem.Name} background removed...");
									}
									catch(Exception ex)
									{
										Console.WriteLine(
											$" Error during conversion: {ex.Message}");
									}
								}
								//Clear(builder);
								//builder.Append("CURL -X POST ");
								//builder.Append(
								//	$"--header \"Ocp-Apim-Subscription-Key:{key}\" ");
								//builder.Append(
								//	"--header \"Content-Type:application/octet-stream\" ");
								//builder.Append("--data-binary \"@");
								//builder.Append(fileInfoItem.FullName);
								//builder.Append("\" ");
								//builder.Append("--output \"");
								//builder.Append(
								//	Path.Combine(item.OutputDir.FullName, fileInfoItem.Name));
								//builder.Append("\" ");
								//builder.Append("\"https://");
								//builder.Append(endpoint);
								//builder.Append(".cognitiveservices.azure.com/");
								//builder.Append("computervision/imageanalysis:segment");
								//builder.Append("?api-version=2023-02-01-preview");
								//builder.Append("&mode=backgroundRemoval");
								//builder.Append("\"");
								//RunExe(@"C:\Windows\System32", "curl.exe",
								//	builder.ToString());
								//Console.WriteLine(
								//	$" {fileInfoItem.Name} background removed...");
							}
							client.Dispose();
						}
						else
						{
							Console.WriteLine(" Error: Output folder must be specified.");
						}
					}
				}
				else
				{
					Console.WriteLine(" Error: Both the\r\n " +
						"MSCognitiveServicesKey and MSCognitiveServicesEndpoint\r\n " +
						"properties must be supplied.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RemoveBackgroundBytes																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Remove background from the current working file.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which variables will be resolved.
		/// </param>
		private static async Task RemoveBackgroundBytes(FileActionItem item)
		{
			byte[] bytes = null;
			HttpClient client = null;
			string endpoint = "";
			Uri hostUri = null;
			string hostUrl = "";
			string key = "";
			HttpResponseMessage response = null;

			if(item != null && WorkingImage?.Bitmap != null)
			{
				key = GetPropertyByName(item, "MSCognitiveServicesKey");
				endpoint = GetPropertyByName(item, "MSCognitiveServicesEndpoint");
				if(key.Length > 0 && endpoint.Length > 0)
				{
					//	The key and the endpoint were both provided.
					client = new HttpClient();
					client.DefaultRequestHeaders.Add(
						"Ocp-Apim-Subscription-Key", key);
					//client.DefaultRequestHeaders.Remove("Content-Type");
					//client.DefaultRequestHeaders.Add(
					//	"Content-Type", "application/octet-stream");
					hostUrl = $"https://{endpoint}.cognitiveservices.azure.com" +
						"/computervision/imageanalysis:segment" +
						"?api-version=2023-02-01-preview" +
						"&mode=backgroundRemoval";
					hostUri = new Uri(hostUrl);
					using(MemoryStream stream = new MemoryStream())
					{
						WorkingImage.Bitmap.Save(stream, ImageFormat.Png);
						bytes = stream.ToArray();
					}
					using(ByteArrayContent content = new ByteArrayContent(bytes))
					{
						//	Check to see if the default headers section works.
						content.Headers.Remove("Content-Type");
						content.Headers.Add(
							"Content-Type", "application/octet-stream");
						try
						{
							response = await client.PostAsync(hostUri, content);
							bytes = await response.Content.ReadAsByteArrayAsync();
							// Use the the following save when not altering the movie.
							using(MemoryStream stream = new MemoryStream(bytes))
							{
								WorkingImage.Bitmap = new Bitmap(stream);
							}
							Console.WriteLine(
								$"  {item.CurrentFile.Name} background removed...");
						}
						catch(Exception ex)
						{
							Console.WriteLine(
								$" Error during conversion: {ex.Message}");
						}
					}
					client.Dispose();
				}
				else
				{
					Console.WriteLine(" Error: Both the\r\n " +
						"MSCognitiveServicesKey and MSCognitiveServicesEndpoint\r\n " +
						"properties must be supplied.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RenumberFiles																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Renumber files to be sequential.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void RenumberFiles(FileActionItem item)
		{
			int count = 0;
			int index = 0;
			FileInfo[] files = null;
			List<FileInfo> filesMatch = null;
			List<string> targetFilenames = null;

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFolderName))
				{
					if(AssignOutputFolder(item))
					{
						files = item.InputDir.GetFiles();
						filesMatch = new List<FileInfo>();
						filesMatch.AddRange(files);
						if(item.Range.StartValue.Length > 0 ||
							item.Range.EndValue.Length > 0)
						{
							FilterFiles(filesMatch, item.Range);
						}
						if(filesMatch.Count > 0)
						{
							UpdateBaseAndDigits(item,
								filesMatch[0].Name, filesMatch.Count, 0);
							//	At this point, we have a base and digit count.
							//	Generate the target filenames.
							targetFilenames = EnumerateFromBase(
								GetPropertyByName(item, nameof(Base)), item.Digits,
								filesMatch.Count);
							if(targetFilenames.Count >= filesMatch.Count)
							{
								count = filesMatch.Count;
								if(!AnyExist(item.OutputDir.FullName, targetFilenames))
								{
									//	The range is clear for renumbering.
									for(index = 0; index < count; index++)
									{
										filesMatch[index].MoveTo(
											Path.Combine(item.OutputDir.FullName,
											targetFilenames[index]));
										Console.WriteLine($" {targetFilenames[index]}");
									}
								}
								else
								{
									//	Target files are in the way.
									Console.Write(" Error: Files exist with the ");
									Console.WriteLine("specified target range...");
								}
							}
						}
						else
						{
							Console.WriteLine(" Warning: No files found in range.");
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RenameFiles																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Rename files with regular expression find and replace pattern.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void RenameFiles(FileActionItem item)
		{
			string find = "";
			List<FileInfo> filesMatch = null;
			string replace = "";

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.InputFolderName |
					ActionElementEnum.OutputFoldername))
				{
					if(AssignOutputFolder(item))
					{
						if(item.Properties.Exists(x => x.Name == "Find"))
						{
							find = item.Properties.First(x => x.Name == "Find").Value;
						}
						if(item.Properties.Exists(x => x.Name == "Replace"))
						{
							replace = item.Properties.First(x => x.Name == "Replace").Value;
						}
						if(find.Length > 0 && replace.Length > 0)
						{
							filesMatch = new List<FileInfo>();
							//	Disconnect from the mothership so we don't risk breaking the
							//	enumeration. Who knows how things are connected these days.
							//	Do you remember the whole FileSystemObject debacle?
							RenameFilesRecursive(item.InputDir, find, replace, item.Recurse);
						}
						else
						{
							if(find.Length == 0)
							{
								Console.WriteLine(" Error: Find property was not specified.");
							}
							if(replace.Length == 0)
							{
								Console.WriteLine(" Error: Replace property not specified.");
							}
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RenameFilesRecursive																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Performs a recursive file renaming operation based upon all of the
		/// command parameters that have already been set.
		/// </summary>
		/// <param name="dir">
		/// Reference to the directory within which the files will be renamed.
		/// </param>
		/// <param name="find">
		/// Regular expression pattern to find.
		/// </param>
		/// <param name="replace">
		/// Regular expression replacement pattern.
		/// </param>
		/// <param name="recurse">
		/// Value indicating whether to recurse into sub-folders.
		/// </param>
		private static void RenameFilesRecursive(DirectoryInfo dir,
			string find, string replace, bool recurse)
		{
			DirectoryInfo[] dirs = null;
			string filename = "";
			List<FileInfo> files = null;
			string originalName = "";

			if(dir != null && find?.Length > 0 && replace != null)
			{
				files = dir.GetFiles().ToList();
				foreach(FileInfo fileInfo in files)
				{
					if(Regex.IsMatch(fileInfo.Name, find))
					{
						try
						{
							originalName = fileInfo.Name;
							filename = Regex.Replace(fileInfo.Name, find, replace);
							File.Move(fileInfo.FullName,
								Path.Combine(dir.FullName,
									filename));
							Console.WriteLine($" {originalName} -> {filename}");
						}
						catch(Exception ex)
						{
							Console.WriteLine(
								$"Error renaming: {originalName}. {ex.Message}");
						}
					}
				}
				if(recurse)
				{
					dirs = dir.GetDirectories();
					if(dirs.Length > 0)
					{
						foreach(DirectoryInfo dirItem in dirs)
						{
							RenameFilesRecursive(dirItem, find, replace, recurse);
						}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RepeatInsertClip																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Copy and insert a clip at the specified location for a specified number
		/// of times.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void RepeatInsertClip(FileActionItem item)
		{
			List<FileInfo> filesMatch = null;
			int repeatCount = 0;
			int repeatIndex = 0;
			List<string> sourceFilenames = null;
			int sourceCount = 0;
			FileInfo sourceFile = null;
			int sourceIndex = 0;
			int targetCount = 0;
			string targetFilename = "";
			List<string> targetFilenames = null;
			int targetIndex = 0;

			if(item != null)
			{
				//	Repeat insert clip.
				if(CheckElements(item,
					ActionElementEnum.InputFolderName |
					ActionElementEnum.Range))
				{
					//	Input folder and range have been supplied.
					if(AssignOutputFolder(item))
					{
						if(item.Count == 0)
						{
							item.Count = 1;
						}
						sourceFilenames = EnumerateRange(item.Range, item.Digits);
						if(sourceFilenames.Count > 0)
						{
							filesMatch = GetFilesInIndexRange(
								AbsolutePath(GetPropertyByName(item, nameof(WorkingPath)),
								GetPropertyByName(item, nameof(InputFolderName))),
								sourceFilenames);
							if(filesMatch.Count > 0)
							{
								UpdateBaseAndDigits(item, filesMatch[0].Name,
									filesMatch.Count,
									GetMaxIndexValue(filesMatch) + 1);
								//	Base and digits have now been expanded as necessary.
								targetFilenames = EnumerateFromBase(
									GetPropertyByName(item, nameof(Base)), item.Digits,
									filesMatch.Count * (int)item.Count);
								if(targetFilenames.Count > 0)
								{
									UpdateBaseAndDigits(item, targetFilenames[^1],
										targetFilenames.Count,
										GetIndexValue(GetPropertyByName(item, nameof(Base))));
								}
								targetFilenames = BufferIndices(targetFilenames, item.Digits);
								//	Make sure we don't cause a collision.
								if(!AnyExist(item.OutputDir.FullName, targetFilenames))
								{
									//	The range is clear for copying.
									repeatCount = (int)item.Count;
									repeatIndex = 0;
									sourceCount = filesMatch.Count;
									targetCount = targetFilenames.Count;
									targetIndex = 0;
									for(repeatIndex = 0; repeatIndex < repeatCount;
										repeatIndex++)
									{
										for(sourceIndex = 0; sourceIndex < sourceCount;
											sourceIndex++, targetIndex++)
										{
											sourceFile = filesMatch[sourceIndex];
											if(sourceFile.Exists &&
												targetIndex < targetFilenames.Count)
											{
												targetFilename = targetFilenames[targetIndex];
												sourceFile.CopyTo(
													Path.Combine(
														item.OutputDir.FullName, targetFilename));
												Console.WriteLine($" {targetFilename}");
											}
										}
									}
								}
								else
								{
									//	Target files are in the way.
									Console.Write(" Error: Files exist with the ");
									Console.WriteLine("specified target range...");
								}
							}
						}
						else
						{
							Console.WriteLine(
								" Error: Range start must be specified on this action.");
						}
					}
					else
					{
						Console.WriteLine(" Error: Could not assign output folder.");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ReplaceGreenscreen																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Remove background from the selected files, outputting them to the
		/// specified target folder.
		/// </summary>
		/// <param name="item">
		/// Reference to the item from which variables will be resolved.
		/// </param>
		private static async Task ReplaceGreenscreen(FileActionItem item)
		{
			Bitmap background = null;
			Bitmap bitmap = null;
			StringBuilder builder = new StringBuilder();
			byte[] bytes = null;
			HttpClient client = null;
			int count = 0;
			string endpoint = "";
			Graphics graphics = null;
			Uri hostUri = null;
			string hostUrl = "";
			int index = 0;
			string key = "";
			HttpResponseMessage response = null;

			if(item != null)
			{
				key = GetPropertyByName(item, "MSCognitiveServicesKey");
				endpoint = GetPropertyByName(item, "MSCognitiveServicesEndpoint");
				if(key.Length > 0 && endpoint.Length > 0)
				{
					//	The key and the endpoint were both provided.
					if(CheckElements(item, ActionElementEnum.Inputs))
					{
						//	After the files have been tested with the Input option, all of
						//	the wildcards will have been resolved.
						//if(item.OutputFolderName.Length == 0)
						//{
						//	//	If no output folder was specified, the target is the
						//	//	working folder.
						//	item.OutputFolderName =
						//		GetPropertyByName(item, nameof(WorkingPath));
						//}
						if(item.OutputFolderName.Length > 0)
						{
							if(!item.OutputDir.Exists)
							{
								item.OutputDir.Create();
							}
							client = new HttpClient();
							client.DefaultRequestHeaders.Add(
								"Ocp-Apim-Subscription-Key", key);
							//client.DefaultRequestHeaders.Remove("Content-Type");
							//client.DefaultRequestHeaders.Add(
							//	"Content-Type", "application/octet-stream");
							hostUrl = $"https://{endpoint}.cognitiveservices.azure.com" +
								"/computervision/imageanalysis:segment" +
								"?api-version=2023-02-01-preview" +
								"&mode=backgroundRemoval";
							hostUri = new Uri(hostUrl);
							foreach(FileInfo fileInfoItem in item.InputFiles)
							{
								bytes = File.ReadAllBytes(fileInfoItem.FullName);
								using(ByteArrayContent content = new ByteArrayContent(bytes))
								{
									//	Check to see if the default headers section works.
									content.Headers.Remove("Content-Type");
									content.Headers.Add(
										"Content-Type", "application/octet-stream");
									try
									{
										response = await client.PostAsync(hostUri, content);
										bytes = await response.Content.ReadAsByteArrayAsync();
										using(MemoryStream stream = new MemoryStream(bytes))
										{
											bitmap = new Bitmap(stream);
											bytes = ImageToColorBytes(bitmap);
											BlankGreenOverRatio(bytes, 255, 102, 128, 102);
											BlankGreenOverRatio(bytes, 255, 102, 143, 107);
											BlankGreenOverRatio(bytes, 255, 110, 140, 105);
											BlankGreenOverRatio(bytes, 255, 105, 143, 99);
											BlankGreenOverRatio(bytes, 255, 105, 135, 105);
											BlankAlphaLevel(bytes, 250);
											bitmap = ColorBytesToImage(bytes, bitmap.PixelFormat,
												bitmap.Width, bitmap.Height);
											//	Create the green background.
											bytes = new byte[bitmap.Width * bitmap.Height * 4];
											count = bytes.Length;
											for(index = 0; index < count; index += 4)
											{
												bytes[index + 1] = bytes[index + 3] = 0xff;
											}
											background = ColorBytesToImage(bytes, bitmap.PixelFormat,
												bitmap.Width, bitmap.Height);
											graphics = Graphics.FromImage(background);
											graphics.CompositingMode = CompositingMode.SourceOver;
											graphics.CompositingQuality =
												CompositingQuality.HighQuality;
											graphics.InterpolationMode =
												InterpolationMode.HighQualityBicubic;
											graphics.SmoothingMode = SmoothingMode.AntiAlias;
											graphics.DrawImage(bitmap, 0, 0);
											graphics.Dispose();
											bitmap = background;
											//	Store the image.
											bitmap.Save(
												Path.Combine(
													item.OutputDir.FullName, fileInfoItem.Name));
										}
										//	Use the the following save when not altering the movie.
										//File.WriteAllBytes(
										//	Path.Combine(item.OutputDir.FullName, fileInfoItem.Name),
										//	bytes);
										Console.WriteLine(
											$" {fileInfoItem.Name} background replaced...");
									}
									catch(Exception ex)
									{
										Console.WriteLine(
											$" Error during conversion: {ex.Message}");
									}
								}
								//Clear(builder);
								//builder.Append("CURL -X POST ");
								//builder.Append(
								//	$"--header \"Ocp-Apim-Subscription-Key:{key}\" ");
								//builder.Append(
								//	"--header \"Content-Type:application/octet-stream\" ");
								//builder.Append("--data-binary \"@");
								//builder.Append(fileInfoItem.FullName);
								//builder.Append("\" ");
								//builder.Append("--output \"");
								//builder.Append(
								//	Path.Combine(item.OutputDir.FullName, fileInfoItem.Name));
								//builder.Append("\" ");
								//builder.Append("\"https://");
								//builder.Append(endpoint);
								//builder.Append(".cognitiveservices.azure.com/");
								//builder.Append("computervision/imageanalysis:segment");
								//builder.Append("?api-version=2023-02-01-preview");
								//builder.Append("&mode=backgroundRemoval");
								//builder.Append("\"");
								//RunExe(@"C:\Windows\System32", "curl.exe",
								//	builder.ToString());
								//Console.WriteLine(
								//	$" {fileInfoItem.Name} background removed...");
							}
							client.Dispose();
						}
						else
						{
							Console.WriteLine(" Error: Output folder must be specified.");
						}
					}
				}
				else
				{
					Console.WriteLine(" Error: Both the\r\n " +
						"MSCognitiveServicesKey and MSCognitiveServicesEndpoint\r\n " +
						"properties must be supplied.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ResizeImage																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Resize the working image to the rectangle specified in the user
		/// properties Width, and Height.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item describing the rectangle to use on
		/// the working image.
		/// </param>
		private static void ResizeImage(FileActionItem item)
		{
			Bitmap bitmap = null;
			int h = 0;
			Rectangle sourceRect = Rectangle.Empty;
			Rectangle targetRect = Rectangle.Empty;
			int w = 0;

			if(item != null && WorkingImage != null)
			{
				//	The item and the working image are both present.
				w = ToInt(GetPropertyByName(item, "Width"));
				h = ToInt(GetPropertyByName(item, "Height"));
				if(w > 0 && h > 0)
				{
					//	Width and height were supplied.
					Console.WriteLine($"  {w}, {h}");
					sourceRect = new Rectangle(0, 0,
						WorkingImage.Bitmap.Width, WorkingImage.Bitmap.Height);
					targetRect = new Rectangle(0, 0, w, h);
					bitmap = new Bitmap(targetRect.Width, targetRect.Height);
					using(Graphics graphics = Graphics.FromImage(bitmap))
					{
						InitializeGraphics(graphics);
						graphics.DrawImage(WorkingImage.Bitmap,
							targetRect, sourceRect, GraphicsUnit.Pixel);
					}
					WorkingImage.Bitmap = bitmap;
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RunActions																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run all of the unmuted actions in the collection.
		/// </summary>
		/// <param name="actions">
		/// Reference to a collection of file actions to run.
		/// </param>
		/// <returns>
		/// Reference to the asynchronous task that was launched.
		/// </returns>
		private static async Task RunActions(FileActionCollection actions)
		{
			if(actions?.Count > 0)
			{
				//actionItems = this.Actions.FindAll(x =>
				//	!x.Options.Exists(y => y.Name.ToLower() == "mute"));
				//foreach(FileActionItem actionItem in actionItems)
				//{
				//	await actionItem.Run();
				//	if(actionItem.Stop)
				//	{
				//		Console.WriteLine("Batch stopped...");
				//		break;
				//	}
				//}
				foreach(FileActionItem actionItem in actions)
				{
					if(!actionItem.Options.Exists(x => x.Name.ToLower() == "mute"))
					{
						await actionItem.Run();
						if(actionItem.Stop)
						{
							Console.WriteLine("Batch stopped...");
							break;
						}
					}
					else
					{
						Console.WriteLine($"Action {actionItem.Action} is muted...");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RunSequence																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the specified sequence.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action that specifies the sequence to run.
		/// </param>
		/// <remarks>
		/// This method loads the Actions collection from the steps found in
		/// the referenced sequence.
		/// </remarks>
		private static async void RunSequence(FileActionItem item)
		{
			string name = "";
			SequenceItem sequence = null;

			if(item != null)
			{
				name = GetPropertyByName(item, "SequenceName");
				if(name?.Length > 0)
				{
					Console.WriteLine($" {name}");
					sequence =
						item.Sequences.FirstOrDefault(x => x.SequenceName == name);
					if(sequence != null)
					{
						//	Copy all of the actions.
						foreach(FileActionItem actionItem in sequence.Actions)
						{
							item.Actions.Add(DeepCopy(actionItem));
						}
						item.Actions.Parent = item;
						//	Run each action.
						await RunActions(item.Actions);
						//foreach(FileActionItem actionitem in item.Actions)
						//{
						//	await actionitem.Run();
						//	if(actionitem.mStop)
						//	{
						//		Console.WriteLine("Sequence stopped...");
						//		item.Stop = true;
						//		break;
						//	}
						//}
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* SetPropertyAbsolutePath																								*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Set the specified property on the item's parent with an absolute path,
		///// given the relative pattern supplied in PropertyName and PropertyValue.
		///// </summary>
		///// <param name="item">
		///// Reference to the action item describing the property to set.
		///// </param>
		//private static void SetPropertyAbsolutePath(FileActionItem item)
		//{
		//	string propertyName = "";
		//	string propertyValue = "";

		//	if(item != null && item.mParent != null && item.mParent.Parent != null)
		//	{
		//		propertyName = GetPropertyByName(item, "PropertyName");
		//		propertyValue = GetPropertyByName(item, "PropertyValue");
		//		SetPropertyValue(item.mParent.Parent, propertyName,
		//			AbsolutePath(
		//				GetPropertyByName(item, nameof(WorkingPath)), propertyValue));
		//	}
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SetPropertyValue																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Set the value of the specified user property on the target item.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item having the property to update.
		/// </param>
		/// <param name="propertyName">
		/// Name of the property to set.
		/// </param>
		/// <param name="propertyValue">
		/// Value to place on the property.
		/// </param>
		private static void SetPropertyValue(FileActionItem item,
			string propertyName, string propertyValue)
		{
			NameValueItem property = null;

			if(item != null && propertyName?.Length > 0)
			{
				property = item.mProperties.FirstOrDefault(x =>
					x.Name == propertyName);
				if(property == null)
				{
					property = new NameValueItem()
					{
						Name = propertyName
					};
					item.mProperties.Add(property);
				}
				property.Value = propertyValue;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SetWorkingImage																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Set the working image to the one specified by the user property
		/// ImageName.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item to be activated.
		/// </param>
		private static void SetWorkingImage(FileActionItem item)
		{
			BitmapInfoItem imageInfo = null;
			string imageName = "";

			WorkingImage = null;
			if(item != null)
			{
				imageName = GetPropertyByName(item, "ImageName");
				if(imageName?.Length > 0)
				{
					//	Image was specified.
					Console.WriteLine($" {imageName}");
					imageInfo = Images.FirstOrDefault(x => x.Name == imageName);
					if(imageInfo != null)
					{
						WorkingImage = imageInfo;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SizeImage																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Scale the working image to a new size to the dimensions found in the
		/// Width and Height user properties.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item describing the new size to use on
		/// the working image.
		/// </param>
		private static void SizeImage(FileActionItem item)
		{
			Bitmap bitmap = null;
			int height = 0;
			Rectangle targetRect = Rectangle.Empty;
			int width = 0;

			if(item != null && WorkingImage != null)
			{
				//	The item and the working image are both present.
				width = ToInt(GetPropertyByName(item, "Width"));
				height = ToInt(GetPropertyByName(item, "Height"));
				if(width > 0 && height > 0)
				{
					//	Dimensions were supplied.
					Console.WriteLine($" {width}, {height}");
					targetRect = new Rectangle(0, 0, width, height);
					bitmap = new Bitmap(targetRect.Width, targetRect.Height);
					using(Graphics graphics = Graphics.FromImage(bitmap))
					{
						InitializeGraphics(graphics);
						graphics.DrawImage(WorkingImage.Bitmap, targetRect);
					}
					WorkingImage.Bitmap = bitmap;
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* StitchFilePatternToMp4																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Stitch files found in the FilePattern property to the output filename
		/// using FFMPEG.
		/// </summary>
		/// <param name="item">
		/// Reference to the item describing properties of the conversion,
		/// including FilePattern and FrameRate, etc.
		/// </param>
		/// <remarks>
		/// <para>
		/// It is assumed that the FFMPEG executable is in your default path.
		/// </para>
		/// <para>
		/// If not specified, the default frame rate is 24.
		/// </para>
		/// </remarks>
		private static void StitchFilePatternToMp4(FileActionItem item)
		{
			string filePattern = "";
			int frameRate = 24;
			string movieFilename = "";
			StringBuilder parameters = new StringBuilder();
			Process process = null;
			string text = "";

			if(item != null)
			{
				text = GetPropertyByName(item, "FrameRate");
				if(text.Length > 0)
				{
					frameRate = ToInt(text);
				}
				filePattern = GetPropertyByName(item, "FilePattern");
				movieFilename = GetPropertyByName(item, "MovieFilename");
				Console.WriteLine($" Creating video {movieFilename}");
				if(filePattern.Length > 0 && movieFilename.Length > 0)
				{
					parameters.Append("-r ");
					parameters.Append(frameRate);
					parameters.Append(" -f image2 -i \"");
					parameters.Append(filePattern);
					parameters.Append("\" -vcodec libx264 -pix_fmt yuv420p \"");
					parameters.Append(movieFilename);
					parameters.Append("\"");
					process = new Process();
					process.StartInfo = new ProcessStartInfo("FFMPEG",
						parameters.ToString())
					{
						UseShellExecute = false
					};
					process.Start();
					process.WaitForExit();
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* SuffixFilenames																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Suffix the specified files with the value of the suffix option.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item being fulfilled.
		/// </param>
		private static void SuffixFilenames(FileActionItem item)
		{
			string extension = "";
			FileOptionItem fileOption = null;
			string name = "";
			string text = "";

			if(item != null)
			{
				if(CheckElements(item,
					ActionElementEnum.Inputs |
					ActionElementEnum.OptionSuffix))
				{
					//	After the files have been tested for input, all of the
					//	wildcards will have been resolved.
					text = "";
					fileOption = GetOptionByName(item, "Suffix");
					if(fileOption != null)
					{
						text = fileOption.Value;
					}
					foreach(FileInfo fileInfoItem in item.InputFiles)
					{
						if(fileInfoItem.Name.Contains('.'))
						{
							extension = Path.GetExtension(fileInfoItem.Name);
							name = $"{LeftOf(fileInfoItem.Name, ".")}{text}{extension}";
						}
						else
						{
							name = $"{fileInfoItem.Name}{text}";
						}
						fileInfoItem.MoveTo(Path.Combine(
							fileInfoItem.Directory.FullName, name));
						Console.WriteLine($" {name}");
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* UpdateBaseAndDigits																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Update the Base and Digits properties from newly known file name and
		/// count information.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item to update.
		/// </param>
		/// <param name="templateFilename">
		/// Previously used template filename.
		/// </param>
		/// <param name="fileCount">
		/// Count of files to be generated.
		/// </param>
		/// <param name="defaultBaseIndex">
		/// The default index to use when the Base property has not been specified.
		/// </param>
		private static void UpdateBaseAndDigits(FileActionItem item,
			string templateFilename, int fileCount, int defaultBaseIndex = 0)
		{
			int seedMax = 0;
			int seedMin = 0;

			if(item != null && templateFilename?.Length > 0 && fileCount > 0)
			{
				if(item.Base.Length == 0)
				{
					if(item.Digits < fileCount.ToString().Length)
					{
						item.Digits = fileCount.ToString().Length;
					}
					item.Digits = Math.Max(GetDigitCount(templateFilename), item.Digits);
					item.Base = FilePatternWithIndex(templateFilename,
						defaultBaseIndex, item.Digits);
				}
				else
				{
					seedMin = GetIndexValue(GetPropertyByName(item, nameof(Base)));
					seedMax = seedMin + fileCount - 1;
					item.Digits = Math.Max(Math.Max(item.Digits,
						GetDigitCount(GetPropertyByName(item, nameof(Base)))),
						seedMax.ToString().Length);
				}
			}
		}
		//*-----------------------------------------------------------------------*

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
		/// Create a new instance of the FileActionItem Item.
		/// </summary>
		public FileActionItem()
		{
			InitializeProperties();
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Action																																*
		//*-----------------------------------------------------------------------*
		private ActionTypeEnum mAction = ActionTypeEnum.None;
		/// <summary>
		/// Get/Set the action associated with this entry.
		/// </summary>
		/// <remarks>
		/// This property is non-inheritable.
		/// </remarks>
		[JsonConverter(typeof(StringEnumConverter))]
		public ActionTypeEnum Action
		{
			get { return mAction; }
			set { mAction = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Actions																																*
		//*-----------------------------------------------------------------------*
		private FileActionCollection mActions = new FileActionCollection();
		/// <summary>
		/// Get a reference to the collection of child file actions.
		/// </summary>
		/// <remarks>
		/// This property is non-inheritable.
		/// </remarks>
		public FileActionCollection Actions
		{
			get { return mActions; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Base																																	*
		//*-----------------------------------------------------------------------*
		private string mBase = null;
		/// <summary>
		/// Get/Set the base number or filename pattern of the source or target
		/// files, depending upon the action.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public string Base
		{
			get
			{
				string result = mBase;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetBase();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mBase = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Bytes																																	*
		//*-----------------------------------------------------------------------*
		private byte[] mBytes = null;
		/// <summary>
		/// Get/Set a reference to a binary byte buffer to use in this session.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		[JsonIgnore]
		public byte[] Bytes
		{
			get
			{
				byte[] result = mBytes;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetBytes();
					}
					else
					{
						result = new byte[0];
					}
				}
				return result;
			}
			set { mBytes = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Conditions																														*
		//*-----------------------------------------------------------------------*
		private ConditionCollection mConditions = new ConditionCollection();
		/// <summary>
		/// Get a reference to the collection of conditions assigned to this
		/// action.
		/// </summary>
		/// <remarks>
		/// This property is not inheritable. However, properties from parent
		/// levels are retrieved when calling the GetConditions function.
		/// </remarks>
		public ConditionCollection Conditions
		{
			get { return mConditions; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ConfigFilename																												*
		//*-----------------------------------------------------------------------*
		private string mConfigFilename = "";
		/// <summary>
		/// Get/Set the path and filename of the configuration file for this
		/// action.
		/// </summary>
		/// <remarks>
		/// This property is non-inheritable.
		/// </remarks>
		public string ConfigFilename
		{
			get { return mConfigFilename; }
			set { mConfigFilename = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Count																																	*
		//*-----------------------------------------------------------------------*
		private float mCount = float.MinValue;
		/// <summary>
		/// Get/Set the count associated with the current action.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public float Count
		{
			get
			{
				float result = mCount;

				if(result == float.MinValue)
				{
					if(mParent != null)
					{
						result = mParent.GetCount();
					}
					else
					{
						result = 0f;
					}
				}
				return result;
			}
			set { mCount = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	CurrentFile																														*
		//*-----------------------------------------------------------------------*
		private FileInfo mCurrentFile = null;
		/// <summary>
		/// Get/Set a reference to the current active file in-use.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		[JsonIgnore]
		public FileInfo CurrentFile
		{
			get
			{
				FileInfo result = mCurrentFile;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetCurrentFile();
					}
				}
				return result;
			}
			set { mCurrentFile = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	CurrentFilename																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Get the currently active filename.
		/// </summary>
		[JsonIgnore]
		public string CurrentFilename
		{
			get
			{
				FileInfo file = GetCurrentFile(this);
				string result = "";

				if(file != null)
				{
					result = file.Name;
				}
				return result;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	CurrentFullFilename																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Get the currently active full path and filename.
		/// </summary>
		[JsonIgnore]
		public string CurrentFullFilename
		{
			get
			{
				FileInfo file = GetCurrentFile(this);
				string result = "";

				if(file != null)
				{
					result = file.FullName;
				}
				return result;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	DateTimeValue																													*
		//*-----------------------------------------------------------------------*
		private DateTime mDateTimeValue = DateTime.MinValue;
		/// <summary>
		/// Get/Set the date and time associated with the current action.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'DateTime'.</para>
		/// </remarks>
		public DateTime DateTimeValue
		{
			get
			{
				DateTime result = mDateTimeValue;

				if(result == DateTime.MinValue && mParent != null)
				{
					result = mParent.GetDateTimeValue();
				}
				return result;
			}
			set { mDateTimeValue = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Digits																																*
		//*-----------------------------------------------------------------------*
		private int mDigits = int.MinValue;
		/// <summary>
		/// Get/Set the number of digits associated with the current action.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public int Digits
		{
			get
			{
				int result = mDigits;

				if(result == int.MinValue)
				{
					if(mParent != null)
					{
						result = mParent.GetDigits();
					}
					else
					{
						result = 0;
					}
				}
				return result;
			}
			set { mDigits = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetConditions																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a collection of conditions defined at the caller's item level
		/// and at all of its parents.
		/// </summary>
		/// <param name="item">
		/// Reference to the file action item to inspect.
		/// </param>
		/// <returns>
		/// Reference to a collection of all conditions defined at the current and
		/// baser levels.
		/// </returns>
		public static ConditionCollection GetConditions(FileActionItem item)
		{
			ConditionCollection conditions = null;
			ConditionCollection result = new ConditionCollection();

			if(item != null)
			{
				if(item.Parent != null && item.Parent.Parent != null)
				{
					conditions = GetConditions(item.Parent.Parent);
					foreach(ConditionItem conditionItem in conditions)
					{
						result.Add(conditionItem);
					}
				}
				//	Write the local items last.
				foreach(ConditionItem conditionItem in item.Conditions)
				{
					result.Add(conditionItem);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetOptionByName																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the option specified by name from this or a parent entity.
		/// </summary>
		/// <param name="item">
		/// Reference to the item for which the option will be found.
		/// </param>
		/// <param name="optionName">
		/// Name of the option to retrieve.
		/// </param>
		/// <returns>
		/// Reference to the specified option, if found. Otherwise, null.
		/// </returns>
		public static FileOptionItem GetOptionByName(FileActionItem item,
			string optionName)
		{
			FileOptionItem result = null;

			if(item != null && optionName?.Length > 0)
			{
				result = item.Options.FirstOrDefault(x =>
					x.Name.ToLower() == optionName.ToLower());
				if(result == null && item.mParent != null)
				{
					result =
						FileActionCollection.GetOptionByName(item.mParent, optionName);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetPropertyByName																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the user property specified by name from this or a parent
		/// entity.
		/// </summary>
		/// <param name="item">
		/// Reference to the item for which the property will be retrieved.
		/// </param>
		/// <param name="propertyName">
		/// Name of the property to retrieve.
		/// </param>
		/// <param name="resolveVariables">
		/// Value indicating whether to resolve variables on this call.
		/// </param>
		/// <returns>
		/// Reference to the specified property, if found. Otherwise, null.
		/// </returns>
		public static string GetPropertyByName(FileActionItem item,
			string propertyName, bool resolveVariables = true)
		{
			PropertyInfo propertySystem = null;
			NameValueItem propertyUser = null;
			object propertyValue = null;
			string result = "";

			if(item != null && propertyName?.Length > 0)
			{
				propertySystem =
					mPublicProperties.FirstOrDefault(x => x.Name.ToLower() ==
					propertyName.ToLower());
				if(propertySystem != null)
				{
					//	Built-in property.
					propertyValue = propertySystem.GetValue(item);
					if(propertyValue != null)
					{
						result = propertyValue.ToString();
					}
				}
				else
				{
					//	User property.
					propertyUser = item.Properties.FirstOrDefault(x =>
						x.Name.ToLower() == propertyName.ToLower());
					if(propertyUser != null)
					{
						result = propertyUser.Value;
					}
					else if(item.mParent != null && item.mParent.Parent != null)
					{
						result = GetPropertyByName(item.mParent.Parent,
							propertyName, false);
					}
				}
				if(result.Length > 0 && resolveVariables)
				{
					result = NormalizeValue(item, result);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Images																																*
		//*-----------------------------------------------------------------------*
		private static BitmapInfoCollection mImages = new BitmapInfoCollection();
		/// <summary>
		/// Get a reference to the collection of images in this session.
		/// </summary>
		[JsonIgnore]
		public static BitmapInfoCollection Images
		{
			get { return mImages; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	InputDir																															*
		//*-----------------------------------------------------------------------*
		private DirectoryInfo mInputDir = null;
		/// <summary>
		/// Get/Set the internal, calculated input directory.
		/// </summary>
		/// <remarks>
		/// This property is non-inerhitable.
		/// </remarks>
		[JsonIgnore]
		public DirectoryInfo InputDir
		{
			get { return mInputDir; }
			set { mInputDir = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	InputFilename																													*
		//*-----------------------------------------------------------------------*
		private string mInputFilename = null;
		/// <summary>
		/// Get/Set the input path and filename of the input file.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'InFile'.</para>
		/// </remarks>
		public string InputFilename
		{
			get
			{
				string result = mInputFilename;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetInputFilename();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mInputFilename = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	InputFiles																														*
		//*-----------------------------------------------------------------------*
		private List<FileInfo> mInputFiles = new List<FileInfo>();
		/// <summary>
		/// Get a reference to the collection of file information used as input in
		/// this session.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		[JsonIgnore]
		public List<FileInfo> InputFiles
		{
			get
			{
				List<FileInfo> result = mInputFiles;

				if(result.Count == 0 && mParent != null)
				{
					result = mParent.GetInputFiles();
				}
				return result;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	InputFolderName																												*
		//*-----------------------------------------------------------------------*
		private string mInputFolderName = null;
		/// <summary>
		/// Get/Set the path and folder name of the input for this action.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'InFolder'.</para>
		/// </remarks>
		public string InputFolderName
		{
			get
			{
				string result = mInputFolderName;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetInputFolderName();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mInputFolderName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	InputNames																														*
		//*-----------------------------------------------------------------------*
		private List<string> mInputNames = new List<string>();
		/// <summary>
		/// Get a reference to the list of filenames or foldernames with
		/// or without wildcards. This parameter can be specified multiple times
		/// on the command line with different values to load multiple input files.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'Inputs'.</para>
		/// </remarks>
		public List<string> InputNames
		{
			get
			{
				List<string> result = mInputNames;

				if(result.Count == 0 && mParent != null)
				{
					//	If the local list is not overridden, then default to the
					//	parent.
					result = mParent.GetInputNames();
				}
				return result;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IsOutputLocal																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the output filenames are local at
		/// this level.
		/// </summary>
		/// <returns>
		/// True if an output filename has been specified at this level.
		/// </returns>
		public bool IsOutputLocal()
		{
			return (mOutputName?.Length > 0);
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Message																																*
		//*-----------------------------------------------------------------------*
		private string mMessage = "";
		/// <summary>
		/// Get/Set a message to be displayed when this action is run.
		/// </summary>
		public string Message
		{
			get { return mMessage; }
			set { mMessage = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Options																																*
		//*-----------------------------------------------------------------------*
		private FileOptionCollection mOptions = new FileOptionCollection();
		/// <summary>
		/// Get a reference to the collection of options assigned to this action.
		/// </summary>
		/// <remarks>
		/// This property is not inheritable. However, options from parent levels
		/// are retrieved when calling the GetOptionByName function.
		/// </remarks>
		public FileOptionCollection Options
		{
			get { return mOptions; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputDir																															*
		//*-----------------------------------------------------------------------*
		private DirectoryInfo mOutputDir = null;
		/// <summary>
		/// Get/Set the internal, calculated output directory.
		/// </summary>
		/// <remarks>
		/// This property is non-inheritable.
		/// </remarks>
		[JsonIgnore]
		public DirectoryInfo OutputDir
		{
			get
			{
				DirectoryInfo directory = mOutputDir;

				if(directory == null && mParent != null && mParent != null)
				{
					directory = mParent.Parent.OutputDir;
				}
				return directory;
			}
			set { mOutputDir = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputFile																														*
		//*-----------------------------------------------------------------------*
		private FileInfo mOutputFile = null;
		/// <summary>
		/// Get/Set the internal, calculated output file.
		/// </summary>
		/// <remarks>
		/// This property is non-inheritable.
		/// </remarks>
		[JsonIgnore]
		public FileInfo OutputFile
		{
			get { return mOutputFile; }
			set { mOutputFile = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputFilename																												*
		//*-----------------------------------------------------------------------*
		private string mOutputFilename = null;
		/// <summary>
		/// Get/Set the output path and filename for this action.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'OutFile'.</para>
		/// </remarks>
		public string OutputFilename
		{
			get
			{
				string result = mOutputFilename;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mOutputFilename;
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mOutputFilename = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputFolderName																											*
		//*-----------------------------------------------------------------------*
		private string mOutputFolderName = null;
		/// <summary>
		/// Get/Set the output path and folder name for this action.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'OutFolder'.</para>
		/// </remarks>
		public string OutputFolderName
		{
			get
			{
				string result = mOutputFolderName;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetOutputFolderName();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mOutputFolderName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputName																														*
		//*-----------------------------------------------------------------------*
		private string mOutputName = null;
		/// <summary>
		/// Get/Set an output pattern that allows for filenames or foldernames
		/// with or without wildcards. This parameter can be specified muliple
		/// times on the command line with different values to write to multiple
		/// output files.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'Output'.</para>
		/// </remarks>
		public string OutputName
		{
			get
			{
				string result = mOutputName;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetOutputName();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mOutputName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	OutputType																														*
		//*-----------------------------------------------------------------------*
		private RenderFileTypeEnum mOutputType = RenderFileTypeEnum.None;
		/// <summary>
		/// Get/Set the type of rendering to be done on the file affected by this
		/// action.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public RenderFileTypeEnum OutputType
		{
			get
			{
				RenderFileTypeEnum result = mOutputType;

				if(result == RenderFileTypeEnum.None)
				{
					if(mParent != null)
					{
						result = mParent.GetOutputType();
					}
					else
					{
						result = RenderFileTypeEnum.Auto;
					}
				}	
				return mOutputType;
			}
			set { mOutputType = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Parent																																*
		//*-----------------------------------------------------------------------*
		private FileActionCollection mParent = null;
		/// <summary>
		/// Get/Set a reference to the parent of this item.
		/// </summary>
		/// <remarks>
		/// This property is non-inheritable.
		/// </remarks>
		[JsonIgnore]
		public FileActionCollection Parent
		{
			get { return mParent; }
			set { mParent = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Pattern																																*
		//*-----------------------------------------------------------------------*
		private string mPattern = null;
		/// <summary>
		/// Get/Set a regular expression pattern for files, folders, or other
		/// appropriate strings.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public string Pattern
		{
			get
			{
				string result = mPattern;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetPattern();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mPattern = value; }
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////*	Prefix																																*
		////*-----------------------------------------------------------------------*
		//private string mPrefix = "";
		///// <summary>
		///// Get/Set a flag indicating that the prefix function is active for the
		///// current action.
		///// </summary>
		///// <remarks>
		///// This property is inheritable.
		///// </remarks>
		//public bool Prefix
		//{
		//	get
		//	{
		//		bool result = ToBool(mPrefix);

		//		if(mPrefix.Length == 0 && mParent != null)
		//		{
		//			result = mParent.GetPrefix();
		//		}
		//		return result;
		//	}
		//	set { mPrefix = value.ToString(); }
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Properties																														*
		//*-----------------------------------------------------------------------*
		private NameValueCollection mProperties = new NameValueCollection();
		/// <summary>
		/// Get a reference to the collection of properties assigned to this
		/// action.
		/// </summary>
		/// <remarks>
		/// This property is not inheritable. However, properties from parent
		/// levels are retrieved when calling the GetPropertyByName function.
		/// </remarks>
		public NameValueCollection Properties
		{
			get { return mProperties; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Range																																	*
		//*-----------------------------------------------------------------------*
		private StartEndItem mRange = null;
		/// <summary>
		/// Get/Set a reference to the start and end values of the range.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public StartEndItem Range
		{
			get
			{
				StartEndItem result = mRange;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetRange();
					}
					else
					{
						result = mRange = new StartEndItem();
					}
				}
				return result;
			}
			set { mRange = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	RectangleInfoList																											*
		//*-----------------------------------------------------------------------*
		private static RectInfoCollection mRectangleInfoList =
			new RectInfoCollection();
		/// <summary>
		/// Get a reference to the collection of rectangle info items in this
		/// session.
		/// </summary>
		[JsonIgnore]
		public static RectInfoCollection RectangleInfoList
		{
			get { return mRectangleInfoList; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Recurse																																*
		//*-----------------------------------------------------------------------*
		private bool mRecurse = false;
		/// <summary>
		/// Get/Set a value indicating whether to recurse folders during the
		/// operation.
		/// </summary>
		/// <remarks>
		/// This property is not inherited.
		/// </remarks>
		public bool Recurse
		{
			get { return mRecurse; }
			set { mRecurse = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Run																																		*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the configured action.
		/// </summary>
		public async Task Run()
		{
			//List<FileActionItem> actionItems = null;
			string content = "";
			FileInfo file = null;
			string lineNumber = "";
			Match match = null;
			string position = "";
			FileActionItem soloItem = null;
			string sourceFilename = "";
			string targetFilename = "";
			FileActionItem topItem = null;

			//	TODO: Create an error exit routine...
			//	Decide which errors require exit and which can just be reported.

			if(mWorkingPathLast != WorkingPath)
			{
				Console.WriteLine($"Working Path: {WorkingPath}");
				mWorkingPathLast = WorkingPath;
			}

			Console.WriteLine($"Action {mAction}...");
			if(Message?.Length > 0)
			{
				Console.WriteLine($" {Message}");
			}

			if(mAction == ActionTypeEnum.Batch)
			{
				//	If this is a batch action, read the contents of
				//	ConfigFilename.
				if(ConfigFilename?.Length > 0)
				{
					sourceFilename = AbsolutePath(
						GetPropertyByName(this, nameof(WorkingPath)),
						GetPropertyByName(this, nameof(ConfigFilename)));
					content = File.ReadAllText(sourceFilename);
					if(content?.Length > 0)
					{
						try
						{
							topItem = JsonConvert.DeserializeObject<FileActionItem>(content);
							if(topItem.Action == ActionTypeEnum.Batch)
							{
								//	All of the top item information is added to this item.
								CopyFields(topItem, this,
									skipList: new string[]
									{
									"mAction", "mConfigFilename",
									"mCurrentFile", "mParent",
									"mWorkingPath"
									});
							}
							else
							{
								//	The top item is a child of this action.
								this.Actions.Add(topItem);
							}
							this.Actions.Parent = this;
							FileActionCollection.InitializeParent(this.Actions);
						}
						catch(Exception ex)
						{
							lineNumber = "Unknown";
							position = "Unknown";
							match = Regex.Match(ex.Message,
								ResourceMain.rxJsonErrorLinePosition);
							if(match.Success)
							{
								lineNumber = GetValue(match, "line");
								position = GetValue(match, "position");
							}
							Console.WriteLine(
								"Error loading configuration file: " +
								$"Line: {lineNumber}, Position: {position}");
						}
					}
					else
					{
						Console.Write("Error: No configuration data loaded from: ");
						Console.WriteLine(sourceFilename);
					}
				}
				else
				{
					Console.WriteLine("Error: Config filename not specified...");
				}
			}
			this.Actions.Parent = this;
			//if(mParent == null)
			//{
			//	//	Initialize all levels from the top level.
			//	InitializeLevels(this);
			//}
			InitializeFilenames(this);
			//if(mAction != ActionTypeEnum.Batch)
			//{
			//	//	When this level isn't a batch, identify all folders and files
			//	//	for the action.
			//	In this version, input files can be defined at any level.
			IdentifyInputFiles(this);
			//}
			IdentifyOutputFiles(this);
			switch(mAction)
			{
				case ActionTypeEnum.AlphaConditionalAdjust:
					// Make adjustments to alpha values of pixels in an image matching
					// the specified values. Available variables are a, r, g, b.
					if(this.CurrentFile == null)
					{
						AlphaConditionalAdjust(this);
					}
					else
					{
						AlphaConditionalAdjustBytes(this);
					}
					break;
				case ActionTypeEnum.AlphaMask:
					AlphaMask(this);
					break;
				case ActionTypeEnum.AntiAliasTransparency:
					//	Smooth the alpha borders between transparent and non-transparent
					//	areas.
					if(this.CurrentFile == null)
					{
						AntiAliasTransparency(this);
					}
					else
					{
						AntiAliasTransparencyBytes(this);
					}
					break;
				case ActionTypeEnum.Batch:
					//	TODO: Allow multiple Soloed items to run.
					//	This is a file batch.
					//	Check first to see if there is a solo.
					soloItem = this.Actions.FirstOrDefault(x =>
						x.Options.Exists(y => y.Name.ToLower() == "solo"));
					if(soloItem != null)
					{
						//	Only run the solo item.
						await soloItem.Run();
					}
					else
					{
						//	Run all non-muted items.
						await RunActions(this.Actions);
						//actionItems = this.Actions.FindAll(x =>
						//	!x.Options.Exists(y => y.Name.ToLower() == "mute"));
						//foreach(FileActionItem actionItem in actionItems)
						//{
						//	await actionItem.Run();
						//	if(actionItem.Stop)
						//	{
						//		Console.WriteLine("Batch stopped...");
						//		break;
						//	}
						//}
					}
					if(IsOutputLocal())
					{
						switch(this.OutputType)
						{
							case RenderFileTypeEnum.RectangleInfoList:
								targetFilename =
									AbsolutePath(
										GetPropertyByName(this, nameof(WorkingPath)),
										GetPropertyByName(this, nameof(OutputName)));
								file = new FileInfo(targetFilename);
								Console.WriteLine(
									$"Writing Rectangles to {file.Name}");
								content = JsonConvert.SerializeObject(RectangleInfoList);
								File.WriteAllText(file.FullName, content);
								break;
						}
					}
					break;
				case ActionTypeEnum.BuildPathProperty:
					BuildPathProperty(this);
					break;
				case ActionTypeEnum.ClearInputFiles:
					ClearInputFiles(this);
					break;
				case ActionTypeEnum.ConvertFromB64:
					//	Convert the file from base-64 to binary.
					ConvertFromB64(this);
					break;
				case ActionTypeEnum.ConvertToB64:
					//	Convert the file from binary to base-64.
					ConvertToB64(this);
					break;
				case ActionTypeEnum.CopyNumericToRange:
					CopyNumericToRange(this);
					break;
				case ActionTypeEnum.CopyRange:
					CopyRange(this);
					break;
				case ActionTypeEnum.CropImage:
					CropImage(this);
					break;
				case ActionTypeEnum.CropImageToRectangleInfoName:
					CropImageToRectangleInfoName(this);
					break;
				case ActionTypeEnum.DelDirectoryPattern:
					DelDirectoryPattern(this);
					break;
				case ActionTypeEnum.DeleteFile:
					DeleteFile(this);
					break;
				case ActionTypeEnum.DelEveryX:
					DelEveryX(this);
					break;
				case ActionTypeEnum.DirReformat:
				case ActionTypeEnum.FormatDirFile:
					//	TODO: Format the DIR file as tab-separated values.
					FormatDirFile(this);
					break;
				case ActionTypeEnum.DirToTsv:
					DirToTsv(this);
					break;
				case ActionTypeEnum.DrawImage:
					DrawImage(this);
					break;
				case ActionTypeEnum.FileOpenImage:
					FileOpenImage(this);
					break;
				case ActionTypeEnum.FileOverlayImage:
					FileOverlayImage(this);
					break;
				case ActionTypeEnum.FileSaveImage:
					FileSaveImage(this);
					break;
				case ActionTypeEnum.FindFiles:
					FindFiles(this);
					break;
				case ActionTypeEnum.ForEachFile:
					ForEachFile(this);
					break;
				case ActionTypeEnum.If:
					//	Run comparisons in this item's Actions collection.
					If(this);
					break;
				case ActionTypeEnum.ImageBackground:
					//	Paint the specified image background color and / or image
					//	on the current working image.
					ImageBackground(this);
					break;
				case ActionTypeEnum.ImagesClear:
					ImagesClear(this);
					break;
				case ActionTypeEnum.ImageSetCommonBoundary:
					ImageSetCommonBoundary(this);
					break;
				case ActionTypeEnum.LoadRectangleInfoList:
					LoadRectangleInfoList(this);
					break;
				case ActionTypeEnum.MoveFiles:
					MoveFiles(this);
					break;
				case ActionTypeEnum.NonLinearEditExcel:
					//	Execute a non-linear editing pattern using an Excel file with
					//	the fields Start, Action, End, Count, X, Y, Width, Height,
					//	and Color.
					//	The input file is expected to be an Excel file.
					await NonLinearEditExcel(this);
					break;
				case ActionTypeEnum.PrefixFilenames:
					PrefixFilenames(this);
					break;
				case ActionTypeEnum.RemoveBackground:
					await RemoveBackground(this);
					break;
				case ActionTypeEnum.RenameFiles:
					RenameFiles(this);
					break;
				case ActionTypeEnum.RenumberFiles:
					RenumberFiles(this);
					break;
				case ActionTypeEnum.RepeatInsertClip:
					RepeatInsertClip(this);
					break;
				case ActionTypeEnum.ReplaceGreenscreen:
					await ReplaceGreenscreen(this);
					break;
				case ActionTypeEnum.RunSequence:
					RunSequence(this);
					break;
				case ActionTypeEnum.SetWorkingImage:
					SetWorkingImage(this);
					break;
				case ActionTypeEnum.SizeImage:
					SizeImage(this);
					break;
				case ActionTypeEnum.StitchFilePatternToMp4:
					StitchFilePatternToMp4(this);
					break;
				case ActionTypeEnum.SuffixFilenames:
					SuffixFilenames(this);
					break;
				default:
					Console.WriteLine($" Error: {Action} not implemented...");
					break;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Sequences																															*
		//*-----------------------------------------------------------------------*
		private SequenceCollection mSequences = new SequenceCollection();
		/// <summary>
		/// Get a reference to the collection of sequences defined for this action.
		/// </summary>
		/// <remarks>
		/// This property is not inheritable.
		/// </remarks>
		public SequenceCollection Sequences
		{
			get
			{
				SequenceCollection result = mSequences;

				if(result.Count == 0 && mParent != null && mParent.Parent != null)
				{
					result = mParent.Parent.Sequences;
				}
				return result;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeActions																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Actions property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeActions()
		{
			return mActions.Count > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeBase																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Base property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeBase()
		{
			return mBase != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeConfigFilename																					*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the ConfigFilename property should
		/// be serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeConfigFilename()
		{
			return mConfigFilename?.Length > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeCount																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Count property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeCount()
		{
			return mCount > float.MinValue;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeDateTimeValue																					*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the DateTimeValue property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeDateTimeValue()
		{
			return mDateTimeValue > DateTime.MinValue;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeDigits																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Digits property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeDigits()
		{
			return mDigits > int.MinValue;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeInputFilename																					*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the InputFilename property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeInputFilename()
		{
			return mInputFilename != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeInputFolderName																				*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the InputFolderName property should
		/// be serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeInputFolderName()
		{
			return mInputFolderName != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeInputNames																							*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the InputNames property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeInputNames()
		{
			return mInputNames.Count > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeOptions																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Options property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeOptions()
		{
			return mOptions.Count > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeOutputFilename																					*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the OutputFilename property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeOutputFilename()
		{
			return mOutputFilename != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeOutputFolderName																				*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the OutputFolderName property should
		/// be serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeOutputFolderName()
		{
			return mOutputFolderName != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeOutputName																							*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the OutputName property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeOutputName()
		{
			return mOutputName != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeOutputType																							*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the OutputType property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeOutputType()
		{
			return mOutputType != RenderFileTypeEnum.None;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializePattern																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Pattern property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializePattern()
		{
			return mPattern != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeProperties																							*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Properties property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeProperties()
		{
			return mProperties.Count > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeRange																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Range property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeRange()
		{
			return mRange != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeSequences																							*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Sequences property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeSequences()
		{
			return mSequences.Count > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeSourceFolderName																				*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the source folder name property
		/// should be serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeSourceFolderName()
		{
			return mSourceFolderName?.Length > 0;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeText																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the Text property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeText()
		{
			return mText != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ShouldSerializeWorkingPath																						*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the WorkingPath property should be
		/// serialized.
		/// </summary>
		/// <returns>
		/// A value indicating whether to serialize the property.
		/// </returns>
		public bool ShouldSerializeWorkingPath()
		{
			return mWorkingPath != null;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	SourceDir																															*
		//*-----------------------------------------------------------------------*
		private DirectoryInfo mSourceDir = null;
		/// <summary>
		/// Get/Set the internal, calculated source directory.
		/// </summary>
		/// <remarks>
		/// This property is non-inerhitable.
		/// </remarks>
		[JsonIgnore]
		public DirectoryInfo SourceDir
		{
			get { return mSourceDir; }
			set { mSourceDir = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	SourceFolderName																											*
		//*-----------------------------------------------------------------------*
		private string mSourceFolderName = null;
		/// <summary>
		/// Get/Set the path and folder name of the data source for this action.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'InFolder'.</para>
		/// </remarks>
		public string SourceFolderName
		{
			get
			{
				string result = mSourceFolderName;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetSourceFolderName();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mSourceFolderName = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Stop																																	*
		//*-----------------------------------------------------------------------*
		private bool mStop = false;
		/// <summary>
		/// Get/Set a value indicating whether the process should be stopped.
		/// </summary>
		[JsonIgnore]
		public bool Stop
		{
			get { return mStop; }
			set
			{
				mStop = value;
				if(mParent?.Parent != null)
				{
					mParent.Parent.Stop = value;
				}
			}
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////*	Suffix																																*
		////*-----------------------------------------------------------------------*
		//private string mSuffix = "";
		///// <summary>
		///// Get/Set a flag indicating that the suffix function is active for the
		///// current action.
		///// </summary>
		///// <remarks>
		///// This property is inheritable.
		///// </remarks>
		//public bool Suffix
		//{
		//	get
		//	{
		//		bool result = ToBool(mSuffix);

		//		if(mSuffix.Length == 0 && mParent != null)
		//		{
		//			result = mParent.GetSuffix();
		//		}
		//		return result;
		//	}
		//	set { mSuffix = value.ToString(); }
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Text																																	*
		//*-----------------------------------------------------------------------*
		private string mText = null;
		/// <summary>
		/// Get/Set the text of the current action.
		/// </summary>
		/// <remarks>
		/// This property is inheritable.
		/// </remarks>
		public string Text
		{
			get
			{
				string result = mText;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetText();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set { mText = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	WorkingImage																													*
		//*-----------------------------------------------------------------------*
		private static BitmapInfoItem mWorkingImage = null;
		/// <summary>
		/// Get/Set a reference to the current working image in this session.
		/// </summary>
		[JsonIgnore]
		public static BitmapInfoItem WorkingImage
		{
			get { return mWorkingImage; }
			set { mWorkingImage = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	WorkingPath																														*
		//*-----------------------------------------------------------------------*
		private string mWorkingPath = null;
		/// <summary>
		/// Get/Set the working path for operations in this instance.
		/// </summary>
		/// <remarks>
		/// <para>This property is inheritable.</para>
		/// <para>Corresponds with the command-line parameter 'Working'.</para>
		/// </remarks>
		public string WorkingPath
		{
			get
			{
				string result = mWorkingPath;

				if(result == null)
				{
					if(mParent != null)
					{
						result = mParent.GetWorkingPath();
					}
					else
					{
						result = "";
					}
				}
				return result;
			}
			set
			{
				mWorkingPath = value;
			}
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}
