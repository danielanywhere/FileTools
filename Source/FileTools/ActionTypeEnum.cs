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
	//*	ActionTypeEnum																													*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Enumeration of available actions.
	/// </summary>
	public enum ActionTypeEnum
	{
		/// <summary>
		/// No action specified or unknown.
		/// </summary>
		None = 0,
		/// <summary>
		/// Make adjustments to alpha values of pixels in an image matching the
		/// specified values. Available variables are a, r, g, b.
		/// </summary>
		AlphaConditionalAdjust,
		/// <summary>
		/// Load a reference file named in InputFilename, and check the red
		/// level of each mask pixel to assure that every apha level of the
		/// working image pixels are at or below the mask level.
		/// </summary>
		AlphaMask,
		/// <summary>
		/// Smooth the alpha borders between transparent and non-transparent areas.
		/// </summary>
		AntiAliasTransparency,
		/// <summary>
		/// Perform a batch of file operations from a single configuration file.
		/// </summary>
		Batch,
		/// <summary>
		/// Build an absolute path in a user property from a partial path template.
		/// </summary>
		BuildPathProperty,
		/// <summary>
		/// Clear the input files collection at this level.
		/// </summary>
		ClearInputFiles,
		/// <summary>
		/// Convert a file from base-64 to binary.
		/// </summary>
		ConvertFromB64,
		/// <summary>
		/// Convert a file from binary to base-64.
		/// </summary>
		ConvertToB64,
		/// <summary>
		/// Copy the first numeric file in the source folder to the target.
		/// </summary>
		CopyFirstNumeric,
		/// <summary>
		/// Copy a range of files in the source folder to a higher increment, in
		/// reverse order.
		/// </summary>
		CopyIncrementReverse,
		/// <summary>
		/// Copy the last numbered file in the source folder to the target.
		/// </summary>
		CopyLastNumeric,
		/// <summary>
		/// Copy the last numbered file in the source folder to an extended number
		/// of frames.
		/// </summary>
		CopyLastNumericExtend,
		/// <summary>
		/// Copy all files from subfolders to the output folder, using merged
		/// sequential naming.
		/// </summary>
		CopyMergeSubs,
		/// <summary>
		/// Copy the specfied numeric frame to the specified range.
		/// </summary>
		CopyNumericToRange,
		/// <summary>
		/// Copy a range of files in the source folder to a new destination.
		/// </summary>
		CopyRange,
		/// <summary>
		/// Crop the working image to the provided Left, Top, Width, and Height
		/// values.
		/// </summary>
		CropImage,
		/// <summary>
		/// Crop the working image to the rectangle specified in the item of the
		/// rectangle list corresponding to the provided Name user property.
		/// </summary>
		CropImageToRectangleInfoName,
		/// <summary>
		/// Copy files or folders from one location to another, adding a date
		/// prefix or suffix to each file at the destination location.
		/// </summary>
		DateCopy,
		/// <summary>
		/// Delete every directory matching the specified pattern.
		/// </summary>
		DelDirectoryPattern,
		/// <summary>
		/// Delete the file specified in the Filename user property if it exists.
		/// </summary>
		DeleteFile,
		/// <summary>
		/// Delete every Xth file in the selected files list.
		/// </summary>
		DelEveryX,
		/// <summary>
		/// Reformat the results of a DIR output file.
		/// </summary>
		DirReformat,
		/// <summary>
		/// Directory to tab-separated values.
		/// </summary>
		DirToTsv,
		/// <summary>
		/// Draw the image specified by ImageName onto the working image at the
		/// location specified by user properties Left and Top.
		/// </summary>
		DrawImage,
		/// <summary>
		/// Open the image file specified in the current input file. Name it
		/// in the local images collection with the name specified in the
		/// user property ImageName.
		/// </summary>
		FileOpenImage,
		/// <summary>
		/// Open each image from the range and place the image specified in
		/// InputFilename at the options specified by Left, Top, Width, and Height.
		/// </summary>
		FileOverlayImage,
		/// <summary>
		/// Save the working image to the currently specified OutputFile.
		/// </summary>
		FileSaveImage,
		/// <summary>
		/// Find the files matching the provided pattern.
		/// </summary>
		FindFiles,
		/// <summary>
		/// Run the Actions collection of the action through all of the files
		/// currently loaded in the InputFiles collection, setting the
		/// CurrentFile property for each pass.
		/// </summary>
		ForEachFile,
		/// <summary>
		/// Format the text output of a DOS DIR * /s command to create a
		/// tab-separated values file.
		/// </summary>
		FormatDirFile,
		/// <summary>
		/// Run one or more conditions to determine whether the sub-actions
		/// of the action should be run.
		/// </summary>
		If,
		/// <summary>
		/// Set the background color of the working image, overlaying the
		/// previous contents on the new background.
		/// </summary>
		ImageBackground,
		/// <summary>
		/// Clear all images from the Images collection.
		/// </summary>
		ImagesClear,
		/// <summary>
		/// Create a list of common boundary box values for sets of images.
		/// </summary>
		ImageSetCommonBoundary,
		/// <summary>
		/// List all files created or modified within the specified date range.
		/// </summary>
		ListAllFilesOnDate,
		/// <summary>
		/// Load the Rectangle info list from an external JSON file.
		/// </summary>
		LoadRectangleInfoList,
		/// <summary>
		/// Move the specified files from one folder to another.
		/// </summary>
		MoveFiles,
		/// <summary>
		/// Execute a non-linear editing pattern using an Excel file with the
		/// fields Start, Action, End, Count, X, Y, Width, Height, Color.
		/// </summary>
		NonLinearEditExcel,
		/// <summary>
		/// Prefix matching filenames with a text pattern.
		/// </summary>
		PrefixFilenames,
		/// <summary>
		/// Remove background from the specified image files, saving the result in
		/// the specified output folder. This action uses Microsoft Azure
		/// Cognitive services and requires the MSCognitiveServicesKey and
		/// MSCognitiveServicesEndpoint properties.
		/// </summary>
		RemoveBackground,
		/// <summary>
		/// Rename matching files with regular expression find and replace.
		/// </summary>
		RenameFiles,
		/// <summary>
		/// Renumber all of the matching files in InputFolder to be contiguous.
		/// </summary>
		RenumberFiles,
		/// <summary>
		/// Repeat a group of one or more files sequentially, renumbering all
		/// following files accordingly.
		/// </summary>
		RepeatInsertClip,
		/// <summary>
		/// Remove background from the specified image files, replacing with a
		/// perfect green screen, saving the result in the specified output
		/// folder. This action uses Microsoft Azure Cognitive services and
		/// requires the MSCognitiveServicesKey and MSCognitiveServicesEndpoint
		/// properties.
		/// </summary>
		ReplaceGreenscreen,
		/// <summary>
		/// Run the sequence specified in the 'SequenceName' user property.
		/// </summary>
		RunSequence,
		/// <summary>
		/// Set the file date and time (TOUCH) on the selected files.
		/// </summary>
		SetFileDate,
		/// <summary>
		/// Set the current working image to the one with the local name
		/// found in the user property ImageName.
		/// </summary>
		SetWorkingImage,
		/// <summary>
		/// Scale the image to a new size, as specified in user properties
		/// Width and Height.
		/// </summary>
		SizeImage,
		/// <summary>
		/// Stitch files found in the FilePattern property to the output filename
		/// using FFMPEG.
		/// </summary>
		StitchFilePatternToMp4,
		/// <summary>
		/// Suffix matching filenames with a text pattern.
		/// </summary>
		SuffixFilenames
	}
	//*-------------------------------------------------------------------------*

}
