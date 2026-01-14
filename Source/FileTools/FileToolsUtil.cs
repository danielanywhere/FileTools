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
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Numerics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Math;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	FileToolsUtil																														*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Utility features and functionality for the FileTools application.
	/// </summary>
	public class FileToolsUtil
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
		//* AbsolutePath																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the absolute path found between the working and relative paths.
		/// </summary>
		/// <param name="workingPath">
		/// The working or default path.
		/// </param>
		/// <param name="relPath">
		/// The relative path or possible fully qualified override.
		/// </param>
		/// <returns>
		/// The absolute path found for the two components.
		/// </returns>
		public static string AbsolutePath(string workingPath, string relPath)
		{
			string result = "";

			if(workingPath?.Length > 0 && (relPath == null || relPath.Length == 0))
			{
				//	Only the working path was specified.
				result = workingPath;
			}
			else if((workingPath == null || workingPath.Length == 0) &&
				relPath?.Length > 0)
			{
				//	Only the relative path was specified.
				result = relPath;
			}
			else if(relPath.Contains(':') || relPath.StartsWith("\\\\") ||
				relPath.StartsWith("//"))
			{
				//	Relative path is a full path.
				result = relPath;
			}
			else
			{
				//	Both the working and relative paths contain information.
				result = Path.Combine(workingPath, relPath);
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AnyExist																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether any of the specified filenames exist
		/// in the specified target folder.
		/// </summary>
		/// <param name="folderName">
		/// Path and folder name of the folder to test.
		/// </param>
		/// <param name="filenames">
		/// Reference to a list of relative path filenames to test for existence.
		/// </param>
		/// <returns>
		/// True if any matches are found. Otherwise, false.
		/// </returns>
		public static bool AnyExist(string folderName, List<string> filenames)
		{
			FileInfo file = null;
			bool result = false;

			if(filenames?.Count > 0)
			{
				foreach(string filename in filenames)
				{
					file = new FileInfo(AbsolutePath(folderName, filename));
					if(file.Exists)
					{
						result = true;
						break;
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* AssureFolder																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Assure the specified folder path exists and return it to the caller if
		/// so.
		/// </summary>
		/// <param name="pathName">
		/// Full path name of the folder to test for.
		/// </param>
		/// <param name="create">
		/// Value indicating whether to create the path if it doesn't yet exist.
		/// </param>
		/// <param name="message">
		/// Message to display with console messages about this folder.
		/// </param>
		/// <param name="quiet">
		/// Value indicating whether to suppress messages.
		/// </param>
		/// <returns>
		/// Reference to the DirectoryInfo representing the folder if it was
		/// possible that the folder existed. Null if the path led to a file
		/// or was not created.
		/// </returns>
		public static DirectoryInfo AssureFolder(string pathName,
			bool create = false, string message = "", bool quiet = false)
		{
			string fullName = "";
			DirectoryInfo result = null;

			if(pathName?.Length > 0)
			{
				fullName = GetFullFoldername(pathName, create, message, quiet);
				if(fullName.Length > 0)
				{
					result = new DirectoryInfo(fullName);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* BlankAlphaLevel																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Blank alpha levels less than or equal to the specified value.
		/// </summary>
		/// <param name="bytes">
		/// Reference to the array of color bytes to inspect, in the arrangement of
		/// BGRA.
		/// </param>
		/// <param name="alpha">
		/// Maximum alpha level to capture.
		/// </param>
		public static void BlankAlphaLevel(byte[] bytes, int alpha)
		{
			byte colorA = 0;
			int count = 0;
			int index = 0;
			byte levelA = 0;

			if(bytes?.Length > 0)
			{
				levelA = (byte)alpha;
				count = bytes.Length;
				for(index = 0; index < count; index += 4)
				{
					colorA = bytes[index + 3];
					if(colorA <= levelA)
					{
						bytes[index] =
							bytes[index + 1] =
							bytes[index + 2] =
							bytes[index + 3] = 0;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* BlankGreenOverRatio																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Blank all pixels over the specified alpha level, less than the
		/// specified red and blue levels, and over the specified green level.
		/// </summary>
		/// <param name="bytes">
		/// Reference to the array of bytes, arranged as BGRA.
		/// </param>
		/// <param name="alpha">
		/// Minimum value of alpha to capture.
		/// </param>
		/// <param name="red">
		/// Maximum value of red to capture.
		/// </param>
		/// <param name="green">
		/// Minimum value od green to capture.
		/// </param>
		/// <param name="blue">
		/// Maximum value of bluw to capture.
		/// </param>
		public static void BlankGreenOverRatio(byte[] bytes, int alpha,
			int red, int green, int blue)
		{
			byte colorA = 0;
			byte colorB = 0;
			byte colorG = 0;
			byte colorR = 0;
			int count = 0;
			int index = 0;
			byte levelA = (byte)alpha;
			byte levelB = (byte)blue;
			byte levelG = (byte)green;
			byte levelR = (byte)red;

			if(bytes?.Length > 0)
			{
				count = bytes.Length;
				for(index = 0; index < count; index += 4)
				{
					colorB = bytes[index];
					colorG = bytes[index + 1];
					colorR = bytes[index + 2];
					colorA = bytes[index + 3];

					if(colorA >= levelA &&
						colorR <= levelR &&
						colorG >= levelG &&
						colorB <= levelB)
					{
						//	Make this item transparent.
						colorA = colorR = colorG = colorB = 0;
						bytes[index] = colorB;
						bytes[index + 1] = colorG;
						bytes[index + 2] = colorR;
						bytes[index + 3] = colorA;
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* BufferIndices																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Buffer the index value in each filename to have at least the specified
		/// number of digits in width.
		/// </summary>
		/// <param name="filenames">
		/// Reference to a list of filenames to inspect.
		/// </param>
		/// <param name="digitCount">
		/// Minimum count of digits in the index.
		/// </param>
		/// <returns>
		/// New list of filenames where each index has been buffered to contain
		/// at least the minimum number of digits specified.
		/// </returns>
		public static List<string> BufferIndices(List<string> filenames,
			int digitCount)
		{
			string firstPart = "";
			string lastPart = "";
			Match match = null;
			List<string> result = new List<string>();
			int seed = 0;

			if(filenames?.Count > 0)
			{
				if(digitCount > 0)
				{
					//	Digits have been specified.
					foreach(string filenameItem in filenames)
					{
						match = Regex.Match(filenameItem, ResourceMain.rxNumericalSeed);
						if(match.Success)
						{
							firstPart = GetValue(match, "pre");
							lastPart = GetValue(match, "post");
							seed = ToInt(GetValue(match, "seed"));
							if(GetValue(match, "seed").Length < digitCount)
							{
								result.Add(
									$"{firstPart}{PadLeft("0", seed, digitCount)}{lastPart}");
							}
							else
							{
								result.Add(filenameItem);
							}
						}
					}
				}
				else
				{
					result.AddRange(filenames);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* Clear																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Clear the contents of the specified string builder.
		/// </summary>
		/// <param name="builder">
		/// Reference to the string builder to be cleared.
		/// </param>
		public static void Clear(StringBuilder builder)
		{
			if(builder?.Length > 0)
			{
				builder.Remove(0, builder.Length);
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ColorBytesToImage																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Convert the caller's binary color bytes to a bitmap image.
		/// </summary>
		/// <param name="bytes">
		/// Reference to the array of color bytes, arranged in BGRA order.
		/// </param>
		/// <param name="pixelFormat">
		/// The pixel format of the image data.
		/// </param>
		/// <param name="width">
		/// Width of the image, in pixels.
		/// </param>
		/// <param name="height">
		/// Height of the image, in pixels.
		/// </param>
		/// <returns>
		/// Reference to the newly created and filled bitmap image representing
		/// the provided color bytes, if legitimate. Otherwise, null;
		/// </returns>
		public static Bitmap ColorBytesToImage(byte[] bytes,
			PixelFormat pixelFormat, int width, int height)
		{
			Bitmap bitmap = null;
			BitmapData bitmapData = null;
			IntPtr ptr;

			if(bytes?.Length > 0 && width > 0 && height > 0)
			{
				using(MemoryStream stream = new MemoryStream(bytes))
				{
					//bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
					bitmap = new Bitmap(width, height, pixelFormat);
					bitmapData = bitmap.LockBits(
						new Rectangle(0, 0, bitmap.Width, bitmap.Height),
						ImageLockMode.WriteOnly, bitmap.PixelFormat);
					ptr = bitmapData.Scan0;
					Marshal.Copy(bytes, 0, ptr, bytes.Length);
					bitmap.UnlockBits(bitmapData);
				}
			}
			return bitmap;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ColorOffset																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the color byte offset for the specified X and Y coordinates.
		/// </summary>
		/// <param name="x">
		/// X coordinate to index.
		/// </param>
		/// <param name="y">
		/// Y coordinate to index.
		/// </param>
		/// <param name="width">
		/// Width of the image, in pixels.
		/// </param>
		/// <param name="colorOff">
		/// Color offset, where 0 - Blue, 1 - Green, 2 - Red, 3 - Alpha.
		/// </param>
		/// <returns>
		/// The absolute offset of the color register for the specified pixel.
		/// </returns>
		public static int ColorOffset(int x, int y, int width, int colorOff = 0)
		{
			int result = (y * width * 4) + (x * 4) + colorOff;
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ConvertWildcardToPattern																							*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Convert the wildcard characters to a regular expression pattern that
		/// can be evaluated for match or no match.
		/// </summary>
		/// <param name="wildcard">
		/// The wildcard pattern to test.
		/// </param>
		/// <returns>
		/// Regular expression pattern that can be used to match values matching
		/// the caller's wildcard.
		/// </returns>
		public static string ConvertWildcardToPattern(string wildcard)
		{
			string result = "";
			if(wildcard?.Length > 0)
			{
				result = wildcard.
					Replace("\\", "\\\\").
					Replace(".", "\\.").
					Replace("*", ".*").
					Replace("?", ".");
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* CopyFields<T>																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Copy the private fields of public properties from the source to target.
		/// </summary>
		/// <typeparam name="T">
		/// Type of object to operate upon.
		/// </typeparam>
		/// <param name="source">
		/// Reference to the source object.
		/// </param>
		/// <param name="target">
		/// Reference to the target object.
		/// </param>
		/// <param name="skipList">
		/// Optional list of field names to skip.
		/// </param>
		public static void CopyFields<T>(T source, T target,
			string[] skipList = null) where T : class
		{
			BindingFlags bindingFlagsF =
				BindingFlags.Instance | BindingFlags.NonPublic;
			BindingFlags bindingFlagsP =
				BindingFlags.Instance | BindingFlags.Public;
			Type elementType = null;
			FieldInfo[] fields = typeof(T).GetFields(bindingFlagsF);
			MethodInfo addMethod = null;
			PropertyInfo[] properties = typeof(T).GetProperties(bindingFlagsP);
			IEnumerable<object> sourceList = null;
			IEnumerable<object> targetList = null;
			object workingValue = null;

			foreach(FieldInfo field in fields)
			{
				if(field.Name.Length > 1 &&
					(skipList == null || !skipList.Contains(field.Name)) &&
					properties.FirstOrDefault(x =>
						x.Name == field.Name.Substring(1)) != null)
				{
					workingValue = field.GetValue(source);
					//if(field.Name == "mOptions")
					//{
					//	Console.WriteLine("CopyFields: Break here...");
					//}
					if(workingValue != null && workingValue is IEnumerable<object>)
					{
						//	The following blind copy is okay, because both lists are
						//	expected to be of the same type.
						sourceList = (IEnumerable<object>)workingValue;
						targetList = (IEnumerable<object>)field.GetValue(target);
						if(sourceList.Count() > 0)
						{
							elementType = sourceList.First().GetType();
							addMethod =
								workingValue.GetType().GetMethod("Add", new[] { elementType });
							foreach(Object item in sourceList)
							{
								addMethod.Invoke(targetList, new object[] { item });
							}
						}
					}
					else
					{
						field.SetValue(target, workingValue);
					}
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* EnumerateFilesAndDirectories																					*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Enumerate through files and directories to return a list of fully
		/// qualified files, given a solid base path and a search pattern that
		/// can include a combination of file and directory names.
		/// </summary>
		/// <param name="directoryPath">
		/// Base directory path from where the search will start.
		/// </param>
		/// <param name="searchPattern">
		/// Search pattern containing wild cards.
		/// </param>
		/// <returns>
		/// Reference to a list of file and folder paths matching the provided
		/// search pattern.
		/// </returns>
		public static List<string> EnumerateFilesAndDirectories(
			string directoryPath, string searchPattern)
		{
			// Replace * and ? characters in search pattern with equivalent regex
			// syntax.
			DirectoryInfo dir = null;
			DirectoryInfo[] dirs = null;
			FileInfo[] files = null;
			int leftPath = 0;
			int leftWild = 0;
			string level = "";
			string regexPattern = "";
			string remainder = "";
			char[] pathMark = new char[] { '\\', '/' };
			List<string> results = new List<string>();
			char[] wild = new char[] { '*', '?' };
			string workingPath = "";
			string workingSearch = "";

			if(directoryPath?.Length > 0 && searchPattern?.Length > 0)
			{
				//	Resolve directories first, then resolve filenames or folders in
				//	each directory.
				//	When entering, the search pattern may continue one or more levels
				//	that are not affected by wildcards. Transfer those to the search
				//	base.
				workingSearch = searchPattern;
				workingPath = directoryPath;
				if(workingSearch.IndexOfAny(wild) > -1 &&
					workingSearch.IndexOfAny(pathMark) > -1)
				{
					//	There are directory marks and wildcards in the search pattern.
					//	Move all non-wild path levels to the left.
					leftWild = workingSearch.IndexOfAny(wild);
					leftPath = workingSearch.IndexOfAny(pathMark);
					while(leftPath > -1 && leftPath < leftWild)
					{
						//	The next character is a path mark. Transfer that to the left.
						workingPath = Path.Combine(workingPath,
							workingSearch.Substring(0, leftPath));
						if(leftPath + 1 < workingSearch.Length)
						{
							workingSearch = workingSearch.Substring(leftPath + 1);
						}
						else
						{
							//	Landing in this location should be impossible.
							workingSearch = "";
							break;
						}
						leftWild = workingSearch.IndexOfAny(wild);
						leftPath = workingSearch.IndexOfAny(pathMark);
					}
					//	At this point, the working path is solid and the working
					//	search contains a wildcard.
					if(workingSearch.IndexOfAny(pathMark) > -1)
					{
						//	Find folder names at the current level that match the
						//	current wildcard.
						level = workingSearch.Substring(0, leftPath);
						if(workingSearch.Length > leftPath + 1)
						{
							remainder = workingSearch.Substring(leftPath + 1);
						}
						else
						{
							remainder = "";
						}
						regexPattern = "^" +
							Regex.Escape(level).
								Replace(@"\*", ".*").
									Replace(@"\?", ".") + "$";
						dir = new DirectoryInfo(workingPath);
						if(dir.Exists)
						{
							//	Directory found.
							dirs = dir.GetDirectories();
							foreach(DirectoryInfo dirItem in dirs)
							{
								if(Regex.IsMatch(dirItem.Name, regexPattern))
								{
									//	This directory is a match.
									if(remainder.Length > 0)
									{
										//	Continue resolving to the right.
										results.AddRange(
											EnumerateFilesAndDirectories(dirItem.FullName,
											remainder));
									}
									else
									{
										//	This is the end of the line for the search.
										//	Most likely, there was a path terminator.
										results.Add(dirItem.FullName);
									}
								}
							}
						}
					}
				}
				//	After base folders have been moved to the directory path,
				//	the search pattern can be resolved.
				if(workingSearch.IndexOfAny(wild) > -1)
				{
					//	There is only an end-level wildcard.
					//	Check for folders and files.
					regexPattern = "^" +
						Regex.Escape(workingSearch).
							Replace(@"\*", ".*").
								Replace(@"\?", ".") + "$";
					dir = new DirectoryInfo(workingPath);
					if(dir.Exists)
					{
						dirs = dir.GetDirectories();
						foreach(DirectoryInfo dirItem in dirs)
						{
							if(Regex.IsMatch(dir.Name, regexPattern))
							{
								results.Add(dirItem.FullName);
							}
						}
						files = dir.GetFiles();
						foreach(FileInfo fileItem in files)
						{
							if(Regex.IsMatch(fileItem.Name, regexPattern))
							{
								results.Add(fileItem.FullName);
							}
						}
					}
				}
				else
				{
					//	Otherwise, the entire search string is the specification.
					results.Add(Path.Combine(workingPath, workingSearch));
				}
			}
			return results;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* EnumerateFromBase																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Enumerate filenames from the specified base.
		/// </summary>
		/// <param name="baseName">
		/// Name of the base filename pattern.
		/// </param>
		/// <param name="digitCount">
		/// Minimum count of digits to apply.
		/// </param>
		/// <param name="fileCount">
		/// Count of files to generate.
		/// </param>
		/// <returns>
		/// Reference to a list of filenames enumerated from the provided base.
		/// </returns>
		public static List<string> EnumerateFromBase(string baseName,
			int digitCount, int fileCount)
		{
			string firstPart = "";
			int index = 0;
			string lastPart = "";
			Match match = null;
			List<string> result = new List<string>();
			int seed = 0;

			if(baseName?.Length > 0 && fileCount > 0)
			{
				match = Regex.Match(baseName, ResourceMain.rxNumericalSeed);
				if(match.Success)
				{
					firstPart = GetValue(match, "pre");
					lastPart = GetValue(match, "post");
					seed = ToInt(GetValue(match, "seed"));
					for(index = 0; index < fileCount; index ++, seed ++)
					{
						result.Add(
							$"{firstPart}{PadLeft("0", seed, digitCount)}{lastPart}");
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	EnumerateRange																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a list of items representing the start through end of the range.
		/// </summary>
		/// <param name="range">
		/// Reference to the range to enumerate.
		/// </param>
		/// <param name="digitCount">
		/// Count of digits required on the output value.
		/// </param>
		/// <param name="defaultExtension">
		/// The default extension to add to the files in the range if one
		/// has not been supplied on the range itself.
		/// </param>
		/// <returns>
		/// Reference to a newly created list of items enumerating all of the
		/// possible items in the range.
		/// </returns>
		/// <remarks>
		/// In this version, a single numerical seed can be surrounded by any
		/// non-numerical value. If more than one numerical seed exist in the
		/// source string, those values will be treated as literal.
		/// If no numerical values are provided, or the pattern doesn't match
		/// between start and end values, only the start and end values are
		/// returned.
		/// </remarks>
		public static List<string> EnumerateRange(StartEndItem range,
			int digitCount = 0, string defaultExtension = "")
		{
			int digits = digitCount;
			string extension = defaultExtension;
			string firstPart = "";
			string lastPart = "";
			int index = 0;
			Match match = null;
			List<string> result = new List<string>();
			int seed1 = 0;
			int seed2 = 0;
			int seedMax = 0;
			int seedMin = 0;

			if(range != null && range.StartValue.Length > 0)
			{
				if(range.StartValue.Contains('.'))
				{
					extension = "";
				}
				else if(defaultExtension?.Length > 0)
				{
					extension = defaultExtension;
				}
				//	If the start value was specified, it will be returned
				//	unconditionally.
				result.Add($"{range.StartValue}{extension}");
				if(range.EndValue.Length > 0)
				{
					//	An end value was specified.
					match = Regex.Match(range.StartValue, ResourceMain.rxNumericalSeed);
					if(match != null && match.Success)
					{
						//	A numerical seed was found in the start.
						firstPart = GetValue(match, "pre");
						lastPart = GetValue(match, "post");
						digits = Math.Max(digits, GetValue(match, "seed").Length);
						seed1 = ToInt(GetValue(match, "seed"));
						match = Regex.Match(range.EndValue, ResourceMain.rxNumericalSeed);
						if(match != null && match.Success)
						{
							seed2 = ToInt(GetValue(match, "seed"));
							digits = Math.Max(digits, GetValue(match, "seed").Length);
							if(firstPart == GetValue(match, "pre") &&
								lastPart == GetValue(match, "post"))
							{
								//	Start and end pattern values align.
								if(seed1 != seed2)
								{
									//	The start and end values refer to different seeds. If
									//	they are equal, only a single item is returned to
									//	the caller.
									seedMin = Math.Min(seed1, seed2);
									seedMax = Math.Max(seed1, seed2);
									for(index = seedMin + 1; index <= seedMax; index ++)
									{
										result.Add(
											$"{firstPart}{PadLeft("0", index, digits)}" +
											$"{lastPart}{extension}");
									}
								}
							}
							else
							{
								//	Start and end pattern values are different.
								result.Add($"{range.EndValue}{extension}");
							}
						}
					}
					else
					{
						//	The first item didn't match a numerical seed so there is no
						//	need to continue.
						result.Add($"{range.EndValue}{extension}");
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FilePatternWithIndex																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a file pattern matching that found in the supplied file, but
		/// having a new index with the specified minimum number of digits in
		/// width.
		/// </summary>
		/// <param name="file">
		/// Reference to the file information item to inspect.
		/// </param>
		/// <param name="index">
		/// The new index value to use in the filename.
		/// </param>
		/// <param name="digitCount">
		/// The minimum count of digits to use in the new filename.
		/// </param>
		/// <returns>
		/// Newly created filename matching the pattern found in the caller's
		/// file, and the newly applied index in the specified minimum number of
		/// digits, padded to the left with zeros as appropriate.
		/// </returns>
		public static string FilePatternWithIndex(FileInfo file, int index,
			int digitCount)
		{
			string firstPart = "";
			string lastPart = "";
			Match match = null;
			string result = "";

			if(file != null)
			{
				match = Regex.Match(file.Name, ResourceMain.rxNumericalSeed);
				if(match.Success)
				{
					firstPart = GetValue(match, "pre");
					lastPart = GetValue(match, "post");
					result = $"{firstPart}{PadLeft("0", index, digitCount)}{lastPart}";
				}
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Return a file pattern matching that found in the supplied file, but
		/// having a new index with the specified minimum number of digits in
		/// width.
		/// </summary>
		/// <param name="filename">
		/// Name of the file to inspect.
		/// </param>
		/// <param name="index">
		/// The new index value to use in the filename.
		/// </param>
		/// <param name="digitCount">
		/// The minimum count of digits to use in the new filename.
		/// </param>
		/// <returns>
		/// Newly created filename matching the pattern found in the caller's
		/// file, and the newly applied index in the specified minimum number of
		/// digits, padded to the left with zeros as appropriate.
		/// </returns>
		public static string FilePatternWithIndex(string filename, int index,
			int digitCount)
		{
			string firstPart = "";
			string lastPart = "";
			Match match = null;
			string result = "";

			if(filename?.Length > 0)
			{
				match = Regex.Match(filename, ResourceMain.rxNumericalSeed);
				if(match.Success)
				{
					firstPart = GetValue(match, "pre");
					lastPart = GetValue(match, "post");
					result = $"{firstPart}{PadLeft("0", index, digitCount)}{lastPart}";
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FilterFiles																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Filter files in the caller-supplied list to only those within the
		/// specified range.
		/// </summary>
		/// <param name="files">
		/// Reference to a collection of files to filter.
		/// </param>
		/// <param name="range">
		/// The range that must be matched by the provided files.
		/// </param>
		public static void FilterFiles(List<FileInfo> files, StartEndItem range)
		{
			int count = 0;
			string firstPart = "";
			string firstPartMax = "";
			string firstPartMin = "";
			int index = 0;
			string lastPart = "";
			string lastPartMax = "";
			string lastPartMin = "";
			Match match = null;
			string name = "";
			int seed = 0;
			int seedMax = 0;
			int seedMin = 0;

			if(files?.Count > 0 && range != null && range.StartValue?.Length > 0 ||
				range.EndValue?.Length > 0)
			{
				if(range.StartValue?.Length > 0)
				{
					match = Regex.Match(range.StartValue, ResourceMain.rxNumericalSeed);
					if(match.Success)
					{
						firstPartMin = GetValue(match, "pre");
						lastPartMin = GetValue(match, "post");
						seedMin = ToInt(GetValue(match, "seed"));
					}
				}
				if(range.EndValue?.Length > 0)
				{
					match = Regex.Match(range.EndValue, ResourceMain.rxNumericalSeed);
					if(match.Success)
					{
						firstPartMax = GetValue(match, "pre");
						lastPartMax = GetValue(match, "post");
						seedMax = ToInt(GetValue(match, "seed"));
					}
				}
				if(firstPartMin == firstPartMax &&
					lastPartMin == lastPartMax)
				{
					//	The range patterns match.
					count = files.Count;
					for(index = 0; index < count; index++)
					{
						name = files[index].Name;
						match = Regex.Match(name, ResourceMain.rxNumericalSeed);
						if(match.Success)
						{
							firstPart = GetValue(match, "pre");
							lastPart = GetValue(match, "post");
							seed = ToInt(GetValue(match, "seed"));
							if(firstPart != firstPartMin ||
								lastPart != lastPartMin)
							{
								//	Pattern doesn't match.
								files.RemoveAt(index);
								index--;		//	Deindex.
								count--;		//	Decount.
							}
							else if(range?.StartValue.Length > 0 &&
								seed < seedMin)
							{
								files.RemoveAt(index);
								index--;    //	Deindex.
								count--;    //	Decount.
							}
							else if(range?.EndValue.Length > 0 &&
								seed > seedMax)
							{
								files.RemoveAt(index);
								index--;    //	Deindex.
								count--;    //	Decount.
							}
						}
					}
				}
				else
				{
					Console.WriteLine(" Error: Range start and end do not match.");
				}
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FindBoundingBox																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the bounding box of the caller's image.
		/// </summary>
		/// <param name="bitmap">
		/// Reference to the bitmap image for which to check the bounding box.
		/// </param>
		/// <returns>
		/// Reference to the bounding box found for the provided image.
		/// </returns>
		/// <remarks>
		/// This version only checks for transparency.
		/// </remarks>
		public static RectInfoItem FindBoundingBox(Bitmap bitmap)
		{
			int byteIndex = 0;
			byte[] bytes = null;
			BitmapData data = null;
			int padding = 0;
			Color pix = Color.Empty;
			int pixelSize = 0;
			RectInfoItem rect = new RectInfoItem();
			int x = 0;
			int xHi = int.MinValue;
			int xLo = int.MaxValue;
			int y = 0;
			int yHi = int.MinValue;
			int yLo = int.MaxValue;

			if(bitmap != null)
			{
				//bitmap.Save(@"C:\Temp\Bitmap.png");
				data = bitmap.LockBits(
					new Rectangle(0, 0, bitmap.Width, bitmap.Height),
					ImageLockMode.ReadOnly, bitmap.PixelFormat);
				pixelSize = (data.PixelFormat == PixelFormat.Format32bppArgb ? 4 : 3);
				padding = data.Stride - (data.Width * pixelSize);
				bytes = new byte[data.Height * data.Stride];
				Console.WriteLine($" Finding bounding box. Pix Size: {pixelSize}...");
				//	Copy the data to the array.
				Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
				bitmap.UnlockBits(data);
				for(y = 0; y < bitmap.Height; y ++)
				{
					for(x = 0; x < bitmap.Width; x ++)
					{
						pix = Color.FromArgb(
							pixelSize == 3 ? 255 : bytes[byteIndex + 3],
							bytes[byteIndex + 2],
							bytes[byteIndex + 1],
							bytes[byteIndex]);
						if(pix.A != 0)
						{
							//	Normal pixel.
							if(x < xLo)
							{
								xLo = x;
							}
							if(x > xHi)
							{
								xHi = x;
							}
							if(y < yLo)
							{
								yLo = y;
							}
							if(y > yHi)
							{
								yHi = y;
							}
						}
						byteIndex += pixelSize;
					}
					byteIndex += padding;
				}
				if(xLo >= 0 && xHi >= xLo && yLo >= 0 && yHi >= yLo)
				{
					rect.Left = xLo - 1;
					rect.Width = xHi - xLo + 2;
					rect.Top = yLo - 1;
					rect.Height = yHi - yLo + 2;
				}
			}
			return rect;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* FindLastDirectoryLevel																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the last static directory level in the caller's directory
		/// expression.
		/// </summary>
		/// <param name="dirExpression">
		/// An abstract directory name pattern containing regular expression
		/// elements.
		/// </param>
		/// <returns>
		/// The portion of the directory to the left of any wildcards,
		/// regular expressions, or interpolated string values.
		/// </returns>
		public static string FindLastDirectoryLevel(string dirExpression)
		{
			int length = 0;
			int location = 0;
			MatchCollection matches = null;
			StringBuilder result = new StringBuilder();

			if(dirExpression?.Length > 0)
			{
				location = dirExpression.IndexOfAny(new char[] { '*', '?', '{' });
				if(location > -1)
				{
					matches = Regex.Matches(dirExpression, @"\\");
					foreach(Match matchItem in matches)
					{
						if(matchItem.Index < location)
						{
							//	The content to the left of the current slash doesn't
							//	contain regular expression material.
							length = matchItem.Index;
						}
						else
						{
							//	We have regular expression materials to the left of the
							//	current slash.
							break;
						}
					}
					if(length > 0)
					{
						//	Return the content to the left of the regular expression level.
						result.Append(dirExpression.Substring(0, length));
					}
				}
				else
				{
					//	There are no regular expression elements in this value.
					result.Append(dirExpression);
				}
			}
			return result.ToString();
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetColumn																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a reference to the column in the table matching a
		/// case-insensitive search for the specified name.
		/// </summary>
		/// <param name="table">
		/// Reference to the data table to inspect.
		/// </param>
		/// <param name="columnName">
		/// Case-insensitive name of the column to find.
		/// </param>
		/// <returns>
		/// Reference to the first column matching the specified name, if found.
		/// Otherwise, null.
		/// </returns>
		public static DataColumn GetColumn(DataTable table, string columnName)
		{
			DataColumn result = null;
			string tl = "";

			if(table != null && table.Columns.Count > 0 && columnName?.Length > 0)
			{
				tl = columnName.ToLower();
				foreach(DataColumn columnItem in table.Columns)
				{
					if(columnItem.ColumnName.ToLower() == tl)
					{
						result = columnItem;
						break;
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetDigitCount																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the active number of digits found in the index portion of the
		/// caller's filename, including left padding.
		/// </summary>
		/// <param name="filename">
		/// Filename to inspect.
		/// </param>
		/// <returns>
		/// Count of digits found in the caller's filename, if found. Otherwise,
		/// 0.
		/// </returns>
		public static int GetDigitCount(string filename)
		{
			Match match = null;
			int result = 0;

			if(filename?.Length > 0)
			{
				match = Regex.Match(filename, ResourceMain.rxNumericalSeed);
				if(match.Success)
				{
					result = GetValue(match, "seed").Length;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetFilesInIndexRange																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return all of the files matching the numerical index range of the
		/// caller's filename list.
		/// </summary>
		/// <param name="path">
		/// Base working path in which to find the files.
		/// </param>
		/// <param name="filenames">
		/// Filenames whose numerical indexes will be matched regardless of
		/// left padding level.
		/// </param>
		/// <returns>
		/// Reference to a list of file information items whose names match the
		/// filename pattern and numerical index values specified in the
		/// caller's specification list.
		/// </returns>
		public static List<FileInfo> GetFilesInIndexRange(string path,
			List<string> filenames)
		{
			DirectoryInfo dir = null;
			FileInfo[] files = null;
			string firstPart = "";
			string lastPart = "";
			List<string> looseNames = null;
			Match match = null;
			List<FileInfo> result = new List<FileInfo>();
			int seed = 0;

			if(path?.Length > 0 && filenames?.Count > 0)
			{
				dir = new DirectoryInfo(path);
				if(dir.Exists)
				{
					files = dir.GetFiles();
					foreach(FileInfo fileInfoItem in files)
					{
						match = Regex.Match(fileInfoItem.Name,
							ResourceMain.rxNumericalSeed);
						if(match.Success)
						{
							firstPart = GetValue(match, "pre");
							lastPart = GetValue(match, "post");
							seed = ToInt(GetValue(match, "seed"));
							looseNames = filenames.FindAll(x =>
								x.StartsWith(firstPart) && x.EndsWith(lastPart));
							if(looseNames.Count > 0)
							{
								foreach(string looseNameItem in looseNames)
								{
									match = Regex.Match(looseNameItem,
										ResourceMain.rxNumericalSeed);
									if(match.Success)
									{
										if(GetValue(match, "pre") == firstPart &&
											GetValue(match, "post") == lastPart &&
											ToInt(GetValue(match, "seed")) == seed)
										{
											//	This file is a match.
											result.Add(fileInfoItem);
											break;
										}
									}
								}
							}
						}
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	GetFullFoldername																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the fully qualified path of the relatively or fully specified
		/// folder.
		/// </summary>
		/// <param name="foldername">
		/// Relative or absolute name of the folder to retrieve.
		/// </param>
		/// <param name="create">
		/// Value indicating whether the folder can be created if it does not
		/// exist.
		/// </param>
		/// <param name="message">
		/// Message to display with folder name.
		/// </param>
		/// <param name="quiet">
		/// Value indicating whether to suppress messages.
		/// </param>
		/// <returns>
		/// Fully qualified path of the specified folder, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public static string GetFullFoldername(string foldername,
			bool create = false, string message = "", bool quiet = false)
		{
			DirectoryInfo dir = null;
			bool exists = false;
			string result = "";

			if(foldername?.Length == 0)
			{
				//	If no folder was specified, use the current working directory.
				dir = new DirectoryInfo(System.Environment.CurrentDirectory);
			}
			else
			{
				//	Some type of filename has been specified.
				if(foldername.StartsWith("\\") || foldername.StartsWith("/") ||
					foldername.IndexOf(":") > -1)
				{
					//	Absolute.
					dir = new DirectoryInfo(foldername);
				}
				else
				{
					//	Relative.
					dir = new DirectoryInfo(
						Path.Combine(System.Environment.CurrentDirectory, foldername));
				}
				exists = dir.Exists;
				if(!exists && !create)
				{
					Console.WriteLine($"Path not found: {message} {dir.FullName}");
					dir = null;
				}
				else if(!exists && create)
				{
					//	Folder can be created.
					dir.Create();
				}
				else if(exists &&
					((dir.Attributes & FileAttributes.Directory) !=
					FileAttributes.Directory))
				{
					//	This object is a file.
					Console.WriteLine("Path is a file. " +
						$"Directory expected: {dir.FullName}");
					dir = null;
				}
			}
			if(dir != null)
			{
				if(!quiet)
				{
					Console.WriteLine($"{message} Directory: {dir.FullName}");
				}
				result = dir.FullName;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetIndexValue																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the binary index value of the numerical filename pattern.
		/// </summary>
		/// <param name="filename">
		/// Name of the file to check for index value.
		/// </param>
		/// <returns>
		/// Numerical index value found within the caller's filename, if found.
		/// Otherwise, 0.
		/// </returns>
		public static int GetIndexValue(string filename)
		{
			Match match = null;
			int result = 0;

			if(filename?.Length > 0)
			{
				match = Regex.Match(filename, ResourceMain.rxNumericalSeed);
				if(match.Success)
				{
					result = ToInt(GetValue(match, "seed"));
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetMaxIndexValue																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the maximum binary index value within the list of numerical
		/// filename patterns.
		/// </summary>
		/// <param name="files">
		/// Reference to a collection of files information items whose indexes
		/// will be checked.
		/// </param>
		/// <returns>
		/// The maximum index value found within the presented filenames, if
		/// found. Otherwise, 0.
		/// </returns>
		public static int GetMaxIndexValue(List<FileInfo> files)
		{
			Match match = null;
			int result = 0;

			if(files?.Count > 0)
			{
				foreach(FileInfo fileItem in files)
				{
					match = Regex.Match(fileItem.Name, ResourceMain.rxNumericalSeed);
					if(match.Success)
					{
						result = Math.Max(result, ToInt(GetValue(match, "seed")));
					}
				}
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Return the maximum binary index value within the list of numerical
		/// filename patterns.
		/// </summary>
		/// <param name="filenames">
		/// Reference to a collection of filenames whose indexes will be checked.
		/// </param>
		/// <returns>
		/// The maximum index value found within the presented filenames, if
		/// found. Otherwise, 0.
		/// </returns>
		public static int GetMaxIndexValue(List<string> filenames)
		{
			Match match = null;
			int result = 0;

			if(filenames?.Count > 0)
			{
				foreach(string filenameItem in filenames)
				{
					match = Regex.Match(filenameItem, ResourceMain.rxNumericalSeed);
					if(match.Success)
					{
						result = Math.Max(result, ToInt(GetValue(match, "seed")));
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	GetRelativeDirectory																									*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the relative portion of the directory name between the new and
		/// base names.
		/// </summary>
		/// <param name="baseName">
		/// Base directory name.
		/// </param>
		/// <param name="newName">
		/// Full name of the sub-directory.
		/// </param>
		/// <returns>
		/// Relative offset name of the two directories.
		/// </returns>
		public static string GetRelativeDirectory(string baseName, string newName)
		{
			int index = 0;
			string result = newName;

			if(baseName?.Length > 0 && newName?.Length > 0 &&
				newName.ToLower().StartsWith(baseName.ToLower()))
			{
				//	The new directory is an extension of the base.
				index = baseName.Length;
				result = newName.Substring(index, newName.Length - index);
				if(result.StartsWith(@"\") || result.StartsWith("/"))
				{
					result = result.Substring(1, result.Length - 1);
				}
				if(result.EndsWith(@"\") || result.EndsWith("/"))
				{
					result = result.Substring(0, result.Length - 1);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetReqNumericDigitCount																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the minimum required numeric digit count required to express
		/// all of the numbers in the given series.
		/// </summary>
		/// <param name="files">
		/// Collection of files whose names will be inspected.
		/// </param>
		/// <returns>
		/// Minimum required number of digits required to support all numeric
		/// values in the collection.
		/// </returns>
		public static int GetReqNumericDigitCount(List<FileInfo> files)
		{
			Match match = null;
			int result = 0;
			int seed = 0;

			if(files?.Count > 0)
			{
				foreach(FileInfo fileInfoItem in files)
				{
					match = Regex.Match(fileInfoItem.Name, ResourceMain.rxNumericalSeed);
					if(match.Success)
					{
						seed = ToInt(GetValue(match, "seed"));
						result = Math.Max(result, seed.ToString().Length);
					}
				}
				if(result == 1)
				{
					result = 0;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* GetValue																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the string value of the specified property within the caller's
		/// File Action.
		/// </summary>
		/// <param name="actionItem">
		/// Reference to the file action item to be inspected.
		/// </param>
		/// <param name="propertyName">
		/// Name of the property to read on the file action action item.
		/// </param>
		/// <returns>
		/// String representation of the specified property value, if found.
		/// Otherwise, an empty string.
		/// </returns>
		public static string GetValue(FileActionItem actionItem,
			string propertyName)
		{
			PropertyInfo property = null;
			string result = "";
			object returned = null;
			Type type = null;

			if(actionItem != null && propertyName?.Length > 0)
			{
				type = actionItem.GetType();
				if(type != null)
				{
					property = type.GetProperty(propertyName);
					if(property != null)
					{
						returned = property.GetValue(actionItem, null);
					}
				}
				if(returned != null)
				{
					result = returned.ToString();
				}
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Return the value of the specified group member in the provided match.
		/// </summary>
		/// <param name="match">
		/// Reference to the match to be inspected.
		/// </param>
		/// <param name="groupName">
		/// Name of the group for which the value will be found.
		/// </param>
		/// <returns>
		/// The value found in the specified group, if found. Otherwise, empty
		/// string.
		/// </returns>
		public static string GetValue(Match match, string groupName)
		{
			string result = "";

			if(match != null && match.Groups[groupName] != null &&
				match.Groups[groupName].Value != null)
			{
				result = match.Groups[groupName].Value;
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Return the value of the specified group member in a match found with
		/// the provided source and pattern.
		/// </summary>
		/// <param name="source">
		/// Source string to search.
		/// </param>
		/// <param name="pattern">
		/// Regular expression pattern to apply.
		/// </param>
		/// <param name="groupName">
		/// Name of the group for which the value will be found.
		/// </param>
		/// <returns>
		/// The value found in the specified group, if found. Otherwise, empty
		/// string.
		/// </returns>
		public static string GetValue(string source, string pattern,
			string groupName)
		{
			Match match = null;
			string result = "";

			if(source?.Length > 0 && pattern?.Length > 0 && groupName?.Length > 0)
			{
				match = Regex.Match(source, pattern);
				if(match.Success)
				{
					result = GetValue(match, groupName);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ImageToColorBytes																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the color byte array of pixel representations of the specified
		/// bitmap.
		/// </summary>
		/// <param name="bitmap">
		/// Reference to the bitmap to convert.
		/// </param>
		/// <returns>
		/// Reference to a byte array containing the color information for each
		/// pixel, in BGRA format, if found. Otherwise, an empty byte array.
		/// </returns>
		public static byte[] ImageToColorBytes(Bitmap bitmap)
		{
			byte[] bytes = new byte[0];
			BitmapData data = null;

			if(bitmap != null)
			{
				data = bitmap.LockBits(
					new Rectangle(Point.Empty, bitmap.Size), ImageLockMode.ReadWrite,
					bitmap.PixelFormat);
				bytes = new byte[data.Height * data.Stride];
				Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
				bitmap.UnlockBits(data);
			}
			return bytes;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IncrementFilename																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Increment the caller's filename and return the next indexed name.
		/// </summary>
		/// <param name="indexedName">
		/// Indexed filename to be incremented.
		/// </param>
		/// <param name="digitCount">
		/// Minimum count of digits in name. If 0, the width is determined
		/// automatically.
		/// </param>
		/// <returns>
		/// Next index within the provided filename pattern, if legal. Otherwise,
		/// an empty string.
		/// </returns>
		public static string IncrementFilename(string indexedName,
			int digitCount = 0)
		{
			int digits = digitCount;
			string firstPart = "";
			string lastPart = "";
			Match match = null;
			string result = "";
			int seed = 0;
			string seedText = "";

			if(indexedName?.Length > 0)
			{
				match = Regex.Match(indexedName, ResourceMain.rxNumericalSeed);
				if(match.Success)
				{
					firstPart = GetValue(match, "pre");
					lastPart = GetValue(match, "post");
					seedText = GetValue(match, "seed");
					seed = ToInt(seedText) + 1;
					digits = Math.Max(seed.ToString().Length, seedText.Length);
					result = $"{firstPart}{PadLeft("0", seed, digits)}{lastPart}";
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* InitializeGraphics																										*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Initialize the active graphics driver with normal settings for this
		/// session.
		/// </summary>
		/// <param name="graphics">
		/// Reference to the graphics device to be initialized.
		/// </param>
		public static void InitializeGraphics(Graphics graphics)
		{
			if(graphics != null)
			{
				graphics.CompositingMode = CompositingMode.SourceOver;
				graphics.CompositingQuality =
					CompositingQuality.HighQuality;
				graphics.SmoothingMode =
					System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
				graphics.InterpolationMode =
					System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IsDirectory																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the specified file is a directory.
		/// </summary>
		/// <param name="fileInfo">
		/// Reference to the file information object to test.
		/// </param>
		/// <returns>
		/// True if the file is a directory. Otherwise, false.
		/// </returns>
		public static bool IsDirectory(FileInfo fileInfo)
		{
			bool result = false;

			if(fileInfo != null && fileInfo.Exists &&
				((int)(fileInfo.Attributes & FileAttributes.Directory)) != 0)
			{
				result = true;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		////*-----------------------------------------------------------------------*
		////* IsNleSourceCut																												*
		////*-----------------------------------------------------------------------*
		///// <summary>
		///// Return a value indicating whether the specified source file index is
		///// in a cut source area.
		///// </summary>
		///// <param name="nle">
		///// The non-linear editing content table.
		///// </param>
		///// <param name="fileIndex">
		///// Source index of the file to test for.
		///// </param>
		///// <returns>
		///// True if the specified index is within a cut range. Otherwise, false.
		///// </returns>
		//public static bool IsNleSourceCut(DataTable nle, int fileIndex)
		//{
		//	//NonLinearEditActionEnum action = NonLinearEditActionEnum.None;
		//	int begin = 0;
		//	int count = 0;
		//	int end = 0;
		//	bool result = false;
		//	string text = "";

		//	if(nle != null && nle.Columns.Contains("Action") &&
		//		nle.Rows.Count > 0 && fileIndex > -1)
		//	{
		//		foreach(DataRow rowItem in nle.Rows)
		//		{
		//			//action = Enum.Parse<NonLinearEditActionEnum>(
		//			//	rowItem.Field<string>("Action"), true);
		//			//switch(action)
		//			//{
		//			//	case NonLinearEditActionEnum.Cut:
		//			//		break;
		//			//	case NonLinearEditActionEnum.FreezeFrame:
		//			//		break;
		//			//	case NonLinearEditActionEnum.MaskRectangle:
		//			//		break;
		//			//}
		//			text = rowItem.Field<string>("Action");
		//			if(text?.Length > 0 && text.ToLower() == "cut")
		//			{
		//				//	A cut specification was found.
		//				begin = ToInt(rowItem.Field<string>("Start"));
		//				end = ToInt(rowItem.Field<string>("End"));
		//				if(end == 0)
		//				{
		//					//	This entry didn't specify an end.
		//					count = ToInt(rowItem.Field<string>("Count"));
		//					if(count > 0)
		//					{
		//						//	A count was specified. Count the number of frames out
		//						//	in the cut.
		//						end = begin + count;
		//					}
		//					else
		//					{
		//						//	No end and no count were specified all indices to the
		//						//	end of the set will be cut.
		//						end = int.MaxValue;
		//					}
		//				}
		//				if(begin > 0 && end > 0)
		//				{
		//					//	Starting and ending indices are present.
		//					if(fileIndex >= begin && fileIndex <= end)
		//					{
		//						result = true;
		//						break;
		//					}
		//				}
		//			}
		//		}
		//	}
		//	return result;
		//}
		////*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IsNumeric																															*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the specified string is totally
		/// numeric.
		/// </summary>
		/// <param name="value">
		/// The value to inspect.
		/// </param>
		/// <param name="floatingPoint">
		/// Value indicating whether to allow floating point values.
		/// </param>
		/// <returns>
		/// True if the caller's value is a valid number. Otherwise, false.
		/// </returns>
		public static bool IsNumeric(string value, bool floatingPoint = true)
		{
			Match match = null;
			bool result = false;

			if(value?.Length > 0)
			{
				if(floatingPoint)
				{
					match = Regex.Match(value, ResourceMain.rxNumericalFloat);
				}
				else
				{
					match = Regex.Match(value, ResourceMain.rxNumericalInt);
				}
				if(match.Success && match.Length == value.Length)
				{
					result = true;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* IsValid																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a value indicating whether the start and end value is valid.
		/// </summary>
		/// <param name="value">
		/// Value indicating whether the start,end range value is valid.
		/// </param>
		public static bool IsValid(StartEndItem value)
		{
			bool result = false;

			result = (value != null &&
				value.StartValue.Length > 0 && value.EndValue.Length > 0);
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* LeftOf																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the portion of the string to the left of the specified pattern.
		/// </summary>
		/// <param name="value">
		/// The original value.
		/// </param>
		/// <param name="pattern">
		/// The pattern at which to stop the original string.
		/// </param>
		/// <param name="last">
		/// Value indicating whether to return the content to the left of the
		/// last instance of the pattern.
		/// </param>
		/// <returns>
		/// Portion of the string to the left of the specified pattern, if
		/// found. Otherwise, the entire value if non-null. Otherwise, an
		/// empty string.
		/// </returns>
		public static string LeftOf(string value, string pattern,
			bool last = false)
		{
			int index = 0;
			string result = "";

			if(value?.Length > 0)
			{
				result = value;
				if(pattern?.Length > 0)
				{
					if(last)
					{
						//	Last index.
						index = value.LastIndexOf(pattern);
					}
					else
					{
						//	First index.
						index = value.IndexOf(pattern);
					}
					if(index > -1)
					{
						result = value.Substring(0, index);
					}
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* PadLeft																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Pad the caller's value to the left with the specified pattern until it
		/// is greater than or equal to the specified total width.
		/// </summary>
		/// <param name="pattern">
		/// Pattern to pad the value with.
		/// </param>
		/// <param name="value">
		/// Value to pad.
		/// </param>
		/// <param name="totalWidth">
		/// The total minimum width allowable.
		/// </param>
		/// <returns>
		/// The caller's value, padded left until it has reached at least the
		/// minimum total width.
		/// </returns>
		public static string PadLeft(string pattern, int value, int totalWidth)
		{
			StringBuilder builder = new StringBuilder();

			builder.Append(value);
			if(pattern?.Length > 0 && totalWidth > 0)
			{
				while(builder.Length < totalWidth)
				{
					builder.Insert(0, pattern);
				}
			}
			return builder.ToString();
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ReadAllImages																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Read all of the specified images to a single bitmap image.
		/// </summary>
		/// <param name="files">
		/// Reference to a collection of file informations.
		/// </param>
		/// <returns>
		/// Bitmap where all of the images in the list have been painted onto
		/// the canvas in the order listed, if successful. Otherwise, a zero
		/// size bitmap.
		/// </returns>
		/// <remarks>
		/// The first file read sets the size for all remaining images in the
		/// list.
		/// </remarks>
		public static Bitmap ReadAllImages(List<FileInfo> files)
		{
			Bitmap bitmap = null;
			int count = 0;
			Graphics graphics = null;
			int index = 0;

			if(files?.Count > 0)
			{
				Console.WriteLine($" Layering {files.Count} images...");
				bitmap = (Bitmap)Bitmap.FromFile(files[0].FullName);
				graphics = Graphics.FromImage(bitmap);
				InitializeGraphics(graphics);

				count = files.Count;
				for(index = 1; index < count; index ++)
				{
					graphics.DrawImage(
						Image.FromFile(files[index].FullName), new Point(0, 0));
				}
				graphics.Dispose();
				graphics = null;
			}
			else
			{
				bitmap = new Bitmap(0, 0);
			}
			return bitmap;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ResolveFilename																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Resolve the supplied path and filename to return all files that match
		/// that name, including wildcards.
		/// </summary>
		/// <param name="fullFilename">
		/// The full filename to parse.
		/// </param>
		/// <param name="create">
		/// Value indicating whether the file will be created if it doesn't yet
		/// exist.
		/// </param>
		/// <returns>
		/// List of existing files found.
		/// </returns>
		/// <remarks>
		/// This method does not distinguish a difference between a file and a
		/// directory. That is left to the calling procedure.
		/// </remarks>
		public static List<FileInfo> ResolveFilename(string fullFilename,
			bool create)
		{
			DirectoryInfo dir = null;
			FileInfo file = null;
			List<string> filenames = null;
			List<FileInfo> files = new List<FileInfo>();
			int leftWild = 0;
			int leftPath = 0;
			char[] pathMark = new char[] { '\\', '/' };
			char[] wild = new char[] { '*', '?' };

			if(fullFilename?.Length > 0)
			{
				if(fullFilename.IndexOfAny(wild) > -1)
				{
					//	The filename contains one or more wildcards. Use regular
					//	expressions to chunk the parts.
					leftPath = fullFilename.IndexOfAny(pathMark);
					leftWild = fullFilename.IndexOfAny(wild);
					if(leftPath > -1 && leftPath < leftWild)
					{
						//	A base path is specified.
						filenames = EnumerateFilesAndDirectories(
							fullFilename.Substring(0, leftPath),
							fullFilename.Substring(leftPath + 1));
					}
					else
					{
						//	The entire path contains wildcards.
						filenames = EnumerateFilesAndDirectories("", fullFilename);
					}
					foreach(string filenameItem in filenames)
					{
						file = new FileInfo(filenameItem);
						if(file.Exists)
						{
							files.Add(file);
						}
					}
				}
				else
				{
					//	No wildcards encountered.
					file = new FileInfo(fullFilename);
					if(!file.Exists)
					{
						//	The file doesn't exist.
						dir = new DirectoryInfo(file.FullName);
						if(dir.Exists)
						{
							//	The file exists and is a directory.
							files.Add(file);
						}
						else if(create)
						{
							files.Add(file);
						}
					}
					else
					{
						//	The file exists.
						files.Add(file);
					}
				}
			}
			return files;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ResolveWildcardFolders																								*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return a list of fully resolved folder names, given a base part and
		/// a wildcard part.
		/// </summary>
		/// <param name="basePath">
		/// Base or static part of the path. This part must be supplied in this
		/// version.
		/// </param>
		/// <param name="wildcardPath">
		/// The wildcard part of the path that can contain one or more levels.
		/// </param>
		/// <returns>
		/// List of fully resolved directories matching the base path and wildcard
		/// pattern.
		/// </returns>
		public static List<string> ResolveWildcardFolders(string basePath,
			string wildcardPath)
		{
			int count = 0;
			DirectoryInfo dir = null;
			DirectoryInfo[] dirs = null;
			string level = "";
			List<string> levels = null;
			string newPath = "";
			string newWildcard = "";
			string pattern = "";
			List<string> results = new List<string>();
			char[] slash = new char[] { '/', '\\' };
			char[] wildcard = new char[] { '*', '?' };

			if(basePath?.Length > 0 && wildcardPath != null)
			{
				if(Directory.Exists(basePath))
				{
					levels = wildcardPath.Split(slash).ToList();
					count = levels.Count;
					if(count > 0)
					{
						//	If there are wildcard levels in the specification, then
						//	resolve from this level inward.
						level = levels[0];
						levels.RemoveAt(0);
						if(levels.Count > 0)
						{
							newWildcard = string.Join('\\', levels);
						}
						else
						{
							newWildcard = "";
						}
						if(level.IndexOfAny(wildcard) > -1)
						{
							//	This level contains one or more wildcards.
							pattern = WildcardToRegEx(level);
							dir = new DirectoryInfo(basePath);
							dirs = dir.GetDirectories();
							foreach(DirectoryInfo dirItem in dirs)
							{
								if(Regex.IsMatch(dirItem.Name, pattern))
								{
									//	This folder matches the wildcard.
									//	Drill inward.
									newPath = Path.Combine(basePath, dirItem.Name);
									if(newWildcard.Length > 0)
									{
										//	Process remaining inner levels.
										results.AddRange(
											ResolveWildcardFolders(newPath, newWildcard));
									}
									else
									{
										//	Working at the last level of the specification.
										results.Add(newPath);
									}
								}
							}
						}
						else
						{
							//	This level is literal.
							newPath = Path.Combine(basePath, level);
							if(Directory.Exists(newPath))
							{
								if(newWildcard.Length > 0)
								{
									//	Process remaining inner levels.
									results.AddRange(
										ResolveWildcardFolders(newPath, newWildcard));
								}
								else
								{
									//	Working at the last level of the specification.
									results.Add(newPath);
								}
							}
						}
					}
					else
					{
						//	The base folder exists.
						//	There are no wildcard levels.
						results.Add(basePath);
					}
				}
				else
				{
					Trace.WriteLine($" Error: Path not found. {basePath}");
				}
			}
			return results;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ResolveWildcards																											*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Resolve any wildcards in the caller's working path and filename.
		/// </summary>
		/// <param name="workingPath">
		/// The base working path to use.
		/// </param>
		/// <param name="filename">
		/// The filename that might contain wildcards.
		/// </param>
		/// <returns>
		/// List of resolved filenames identified.
		/// </returns>
		public static List<string> ResolveWildcards(string workingPath,
			string filename)
		{
			char[] backslashes = new char[] { '\\', '/' };
			StringBuilder builder = new StringBuilder();
			int count = 0;
			string endLevel = "";
			FileInfo[] files = null;
			List<string> filenames = new List<string>();
			string fullFilename = "";
			int index = 0;
			string[] levels = null;
			char[] wildcards = new char[] { '*', '?' };
			string workingLevel = "";

			if(filename?.Length > 0)
			{
				fullFilename = AbsolutePath(workingPath, filename);
				levels = fullFilename.Split(backslashes);
				if(levels.Length > 1)
				{
					//	There is at least a base directory and a file.
					endLevel = levels[^1];
					if(endLevel.IndexOfAny(wildcards) > -1)
					{
						//	The end level contains wildcards.
						count = levels.Length;
						for(index = 0; index < count - 1; index ++)
						{
							if(builder.Length > 0)
							{
								builder.Append('\\');
							}
							builder.Append(levels[index]);
						}
						workingLevel = builder.ToString();
						if(Directory.Exists(workingLevel))
						{
							files = new DirectoryInfo(workingLevel).GetFiles(endLevel);
							foreach(FileInfo fileItem in files)
							{
								filenames.Add(fileItem.FullName);
							}
						}
					}
					else
					{
						//	Use the full filename to get the path.
						if(File.Exists(fullFilename) || Directory.Exists(fullFilename))
						{
							filenames.Add(fullFilename);
						}
					}
				}
			}
			return filenames;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* Right																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the specified length of characters from the right side of the
		/// provided string.
		/// </summary>
		/// <param name="source">
		/// The string to inspect.
		/// </param>
		/// <param name="length">
		/// The length of characters to return.
		/// </param>
		/// <returns>
		/// The rightmost number of characters from the caller's string specified
		/// in the length parameter, if available, otherwise, the full content of
		/// the string, if provided. Otherwise, an empty string.
		/// </returns>
		public static string Right(string source, int length)
		{
			string result = "";

			if(source?.Length > 0 && length > 0)
			{
				if(length >= source.Length)
				{
					result = source;
				}
				else
				{
					//	The return value is narrower than the source.
					result = source.Substring(source.Length - length, length);
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RightOf																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the string portion to the right of the specified pattern.
		/// </summary>
		/// <param name="source">
		/// Source value to inspect.
		/// </param>
		/// <param name="pattern">
		/// The pattern to test for.
		/// </param>
		/// <returns>
		/// The portion of the supplied string to the right of the specified
		/// pattern.
		/// </returns>
		public static string RightOf(string source, string pattern)
		{
			int position = 0;
			string result = "";

			if(source?.Length > 0)
			{
				if(pattern?.Length > 0 && source.IndexOf(pattern) > -1)
				{
					//	The pattern exists in the string.
					position = source.LastIndexOf(pattern);
					if(source.Length > position + 1)
					{
						result = source.Substring(position + 1);
					}
				}
				else
				{
					result = source;
				}
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RunExe																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the EXE file and wait for completion.
		/// </summary>
		/// <param name="exePath">
		/// Path to the executable file.
		/// </param>
		/// <param name="exeName">
		/// Executable name.
		/// </param>
		/// <param name="arguments">
		/// Optional arguments to add to the command line.
		/// </param>
		public static void RunExe(string exePath, string exeName,
			string arguments = "")
		{
			Process process = null;

			if(exePath?.Length > 0 && exeName?.Length > 0)
			{
				process = new Process();
				process.StartInfo.FileName = Path.Combine(exePath, exeName);
				if(arguments?.Length > 0)
				{
					process.StartInfo.Arguments = arguments;
				}

				process.StartInfo.UseShellExecute = true;
				process.Start();
				process.WaitForExit();
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* RunExeConsole																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the EXE file with output directly to the Trace.
		/// </summary>
		/// <param name="exePath">
		/// Path to the executable file.
		/// </param>
		/// <param name="exeName">
		/// Executable name.
		/// </param>
		/// <param name="arguments">
		/// Optional arguments to add to the command line.
		/// </param>
		public static void RunExeConsole(string exePath, string exeName,
			string arguments = "")
		{
			Process process = null;

			if(exePath?.Length > 0 && exeName?.Length > 0)
			{
				process = new Process();
				process.StartInfo.RedirectStandardOutput = true;
				process.StartInfo.RedirectStandardError = true;
				process.StartInfo.FileName = Path.Combine(exePath, exeName);
				if(arguments?.Length > 0)
				{
					process.StartInfo.Arguments = arguments;
				}

				process.StartInfo.UseShellExecute = false;
				process.StartInfo.CreateNoWindow = true;
				process.OutputDataReceived +=
					(sender, args) => Console.WriteLine("  {0}", args.Data);
				process.ErrorDataReceived +=
					(sender, args) => Console.WriteLine("  {0}", args.Data);
				process.Start();
				process.BeginOutputReadLine();
				process.BeginErrorReadLine();
				process.WaitForExit();
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ToBool																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Provide fail-safe conversion of string to boolean value.
		/// </summary>
		/// <param name="value">
		/// Value to convert.
		/// </param>
		/// <returns>
		/// Boolean value. False if not convertible.
		/// </returns>
		public static bool ToBool(object value)
		{
			bool result = false;
			if(value != null)
			{
				result = ToBool(value.ToString());
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Provide fail-safe conversion of string to boolean value.
		/// </summary>
		/// <param name="value">
		/// Value to convert.
		/// </param>
		/// <param name="defaultValue">
		/// The default value to return if the value was not present.
		/// </param>
		/// <returns>
		/// Boolean value. False if not convertible.
		/// </returns>
		public static bool ToBool(string value, bool defaultValue = false)
		{
			//	A try .. catch block was originally implemented here, but the
			//	following text was being sent to output on each unsuccessful
			//	match.
			//	Exception thrown: 'System.FormatException' in mscorlib.dll
			bool result = false;

			if(value?.Length > 0)
			{
				if(!bool.TryParse(value, out result))
				{
					Debug.WriteLine($"Error on ToBool");
				}
			}
			else
			{
				result = defaultValue;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ToDateTime																														*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Parse a string to a date and time value.
		/// </summary>
		/// <param name="value">
		/// The date and time text value to parse.
		/// </param>
		/// <returns>
		/// The specified date and time, represented by the DateTime object, if
		/// valid. Otherwise, the minimum value for a DateTime object.
		/// </returns>
		/// <remarks>
		/// <para>Allowable formats are the following:</para>
		/// <list type="bullet">
		/// <item><b>MM/DD/YYYY[ HH:MM[:SS[.FFF]]]</b>. General US short
		/// format.</item>
		/// <item><b>YYYYMMDD[.HHMM[SS[FFF]]]</b>. Sortable numeric format.</item>
		/// </list>
		/// </remarks>
		public static DateTime ToDateTime(string value)
		{
			int day = 0;
			int hour = 0;
			Match match = null;
			int millisecond = 0;
			int minute = 0;
			int month = 0;
			DateTime now = DateTime.Now;
			DateTime result = DateTime.MinValue;
			int second = 0;
			string text = "";
			int year = 0;

			if(value?.Length > 0)
			{
				//	Value supplied.
				if(value.Contains('/'))
				{
					//	MM/DD/YYYY.
					match = Regex.Match(value, ResourceMain.rxDateMDY);
				}
				else
				{
					//	YYYYMMDD.
					match = Regex.Match(value, ResourceMain.rxDateYYYYMMDD);
				}
				if(match.Success)
				{
					//	A match was found.
					text = GetValue(match, "year");
					if(text.Length == 0)
					{
						text = now.Year.ToString();
					}
					else if(text.Length == 2)
					{
						text = $"20{text}";
					}
					year = ToInt(text);
					text = GetValue(match, "month");
					if(text.Length == 0)
					{
						text = now.Month.ToString();
					}
					month = ToInt(text);
					text = GetValue(match, "day");
					if(text.Length == 0)
					{
						text = now.Day.ToString();
					}
					day = ToInt(text);
					minute = ToInt(GetValue(match, "minute"));
					text = GetValue(match, "second");
					second = ToInt(GetValue(match, "second"));
					millisecond = ToInt(GetValue(match, "millisecond"));
					result =
						new DateTime(year, month, day, hour, minute, second, millisecond);
				}
			}
			return DateTime.MinValue;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ToFloat																																*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Provide fail-safe conversion of string to numeric value.
		/// </summary>
		/// <param name="value">
		/// Value to convert.
		/// </param>
		/// <returns>
		/// Floating point value. 0 if not convertible.
		/// </returns>
		public static float ToFloat(object value)
		{
			float result = 0f;
			if(value != null)
			{
				result = ToFloat(value.ToString());
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Provide fail-safe conversion of string to numeric value.
		/// </summary>
		/// <param name="value">
		/// Value to convert.
		/// </param>
		/// <returns>
		/// Floating point value. 0 if not convertible.
		/// </returns>
		public static float ToFloat(string value)
		{
			float result = 0f;
			try
			{
				result = float.Parse(value);
			}
			catch { }
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ToInt																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Provide fail-safe conversion of string to numeric value.
		/// </summary>
		/// <param name="value">
		/// Value to convert.
		/// </param>
		/// <returns>
		/// Int32 value. 0 if not convertible.
		/// </returns>
		public static int ToInt(object value)
		{
			int result = 0;
			if(value != null)
			{
				result = ToInt(value.ToString());
			}
			return result;
		}
		//*- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -*
		/// <summary>
		/// Provide fail-safe conversion of string to numeric value.
		/// </summary>
		/// <param name="value">
		/// Value to convert.
		/// </param>
		/// <returns>
		/// Int32 value. 0 if not convertible.
		/// </returns>
		public static int ToInt(string value)
		{
			int result = 0;
			try
			{
				result = int.Parse(value);
			}
			catch { }
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* ValidatePath																													*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Return the fully qualified path of the folder or file if it exists or
		/// can be created.
		/// </summary>
		/// <param name="pathType">
		/// Name of the type of path to inspect.
		/// </param>
		/// <param name="workingPath">
		/// The working path to reference if the main path was relative.
		/// </param>
		/// <param name="relativePath">
		/// The relative or absolute path to validate.
		/// </param>
		/// <param name="canCreate">
		/// Value indicating whether the folder or file can be created.
		/// </param>
		/// <returns>
		/// If the path or filename was legitimate, the fully qualified name is
		/// returned. Otherwise, an empty string is returned.
		/// </returns>
		public static string ValidatePath(string pathType,
			string workingPath, string relativePath,
			bool canCreate = false)
		{
			string path = "";
			string result = "";

			if(relativePath?.Length > 0)
			{
				path = Path.GetFullPath(AbsolutePath(workingPath, relativePath));
				if(path.Length > 0 &&
					(Path.Exists(path) || canCreate))
				{
					result = path;
				}
				else if(path.Length == 0)
				{
					Trace.WriteLine($" {pathType} Error: No path was specified.");
				}
				else
				{
					Trace.WriteLine($" {pathType} Error: Path or file not found...");
				}
			}
			else if(workingPath?.Length > 0)
			{
				result = workingPath;
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//* WildcardToRegEx																												*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Convert the caller's wildcard pattern to a regular expression pattern
		/// and return that to the caller.
		/// </summary>
		/// <param name="wildcardPattern">
		/// The wildcard pattern to convert, potentially containing '*' and '?'
		/// characters.
		/// </param>
		/// <returns>
		/// The wildcard pattern, usable as a regular expression match pattern.
		/// </returns>
		public static string WildcardToRegEx(string wildcardPattern)
		{
			string result = "";

			if(wildcardPattern?.Length > 0)
			{
				result = wildcardPattern.Replace("?", "(.{0,1})");
				result = result.Replace("*", "(.*?)");
				result = $"^{result}$";
			}
			return result;
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}
