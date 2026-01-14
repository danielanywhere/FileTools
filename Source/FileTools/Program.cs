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
using System.IO;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;

using static FileTools.FileToolsUtil;

namespace FileTools
{
	//*-------------------------------------------------------------------------*
	//*	Program																																	*
	//*-------------------------------------------------------------------------*
	/// <summary>
	/// Main instance of the FileTools application.
	/// </summary>
	public class Program
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
		//*	_Main																																	*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Configure and run the application.
		/// </summary>
		public static async Task Main(string[] args)
		{
			//	20230425.1414 - Prefix and Suffix are how handled in options.
			ActionTypeEnum action = ActionTypeEnum.None;
			bool bShowCommand = false;
			bool bShowHelp = false; //	Flag - Explicit Show Help.
			StringBuilder builder = new StringBuilder();
			char[] comma = new char[] { ',' };
			string key = "";        //	Current Parameter Key.
			string lowerArg = "";   //	Current Lowercase Argument.
															//	Message to display in Console.
			StringBuilder message = new StringBuilder();
			NameValueCollection namevalues = null;
			Program prg = new Program();  //	Initialized instance.
			string[] values = null;

			Console.WriteLine("FileTools.exe");

			//	NOTE: Wildcards still don't work on paths.
			//Console.WriteLine("Test wildcard with path...");
			//string[] paths = Directory.GetDirectories(@"C:\Temp\Movies\*st*");
			//foreach(string pathItem in paths)
			//{
			//	string[] filenames = Directory.GetFileSystemEntries($"{pathItem}\\ReadMe.txt");
			//	foreach(string filenameItem in filenames)
			//	{
			//		Console.WriteLine(filenameItem);
			//	}
			//}
			//return;

			foreach(string arg in args)
			{
				lowerArg = arg.ToLower();
				key = "/?";
				if(lowerArg == key)
				{
					bShowHelp = true;
					continue;
				}
				key = "/action:";
				if(lowerArg.StartsWith(key))
				{
					if(Enum.TryParse<ActionTypeEnum>(arg.Substring(key.Length),
						true, out action))
					{
						if(action != ActionTypeEnum.None)
						{
							prg.ActionItem.Action = action;
						}
						else
						{
							message.Append("Error: No action specified...");
							bShowHelp = true;
							break;
						}
					}
					continue;
				}
				key = "/base:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Base = arg.Substring(key.Length);
					continue;
				}
				key = "/commandline";
				if(lowerArg == key)
				{
					bShowCommand = true;
					continue;
				}
				key = "/configfile:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.ConfigFilename = arg.Substring(key.Length);
					continue;
				}
				key = "/count:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Count = ToFloat(arg.Substring(key.Length));
					continue;
				}
				key = "/datetime:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.DateTimeValue = ToDateTime(arg.Substring(key.Length));
					continue;
				}
				key = "/digits:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Digits = ToInt(arg.Substring(key.Length));
					continue;
				}
				//key = "/exampleparameter:";
				//if(lowerArg.StartsWith(key))
				//{
				//	prg.exampleparameter = arg.Substring(key.Length);
				//	continue;
				//}
				key = "/infile:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.InputFilename = arg.Substring(key.Length);
					continue;
				}
				key = "/infolder:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.InputFolderName = arg.Substring(key.Length);
					continue;
				}
				key = "/input:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.InputNames.Add(arg.Substring(key.Length));
					continue;
				}
				key = "/option:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Options.Add(arg.Substring(key.Length));
					continue;
				}
				key = "/outfile:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.OutputFilename = arg.Substring(key.Length);
					continue;
				}
				key = "/outfolder:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.OutputFolderName = arg.Substring(key.Length);
					continue;
				}
				key = "/output:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.OutputName = arg.Substring(key.Length);
					continue;
				}
				key = "/pattern:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Pattern = arg.Substring(key.Length);
					continue;
				}
				//key = "/prefix";
				//if(lowerArg == key)
				//{
				//	prg.ActionItem.Prefix = true;
				//}
				key = "/properties:";
				if(lowerArg.StartsWith(key))
				{
					try
					{
						namevalues = JsonConvert.DeserializeObject<NameValueCollection>(
							arg.Substring(key.Length));
						foreach(NameValueItem propertyItem in namevalues)
						{
							prg.mActionItem.Properties.Add(propertyItem);
						}
					}
					catch(Exception ex)
					{
						Console.WriteLine($"Error parsing properties: {ex.Message}");
						bShowHelp = true;
					}
					continue;
				}
				key = "/range:";
				if(lowerArg.StartsWith(key))
				{
					values = arg.Substring(key.Length).Split(comma);
					if(values.Length > 0)
					{
						prg.ActionItem.Range.StartValue = values[0];
					}
					if(values.Length > 1)
					{
						prg.ActionItem.Range.EndValue = values[1];
					}
					continue;
				}
				key = "/recurse";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Recurse = true;
					continue;
				}
				//key = "/suffix";
				//if(lowerArg == key)
				//{
				//	prg.ActionItem.Suffix = true;
				//}
				key = "/text:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.Text = arg.Substring(key.Length);
					continue;
				}
				key = "/wait";
				if(lowerArg.StartsWith(key))
				{
					prg.mWaitAfterEnd = true;
					continue;
				}
				key = "/workingpath:";
				if(lowerArg.StartsWith(key))
				{
					prg.ActionItem.WorkingPath = arg.Substring(key.Length);
					continue;
				}
				//	Compatibility with earlier version. Remove at first opportunity.
				key = "/working:";
				if(lowerArg.StartsWith(key))
				{
					Console.WriteLine("WARNING: The /working: parameter is obsolete.");
					Console.WriteLine(
						" Please use the /workingpath: parameter instead.");
					prg.ActionItem.WorkingPath = arg.Substring(key.Length);
					continue;
				}
			}
			if(prg.ActionItem.WorkingPath.Length == 0)
			{
				//	Working path has not been specified. Use the current operating
				//	path.
				prg.ActionItem.WorkingPath = Directory.GetCurrentDirectory();
			}
			//Console.WriteLine($"Working Path: {prg.ActionItem.WorkingPath}");
			if(bShowCommand)
			{
				Clear(builder);
				builder.Append("Command:");
				foreach(string argItem in args)
				{
					builder.Append(' ');
					builder.Append(argItem);
				}
			}
			if(bShowHelp)
			{
				//	Display Syntax.
				Console.WriteLine(message.ToString() + "\r\n" + ResourceMain.Syntax);
			}
			else
			{
				//	Run the configured application.
				await prg.Run();
			}
			if(prg.mWaitAfterEnd)
			{
				Console.WriteLine("Press [Enter] to exit...");
				Console.ReadLine();
			}
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	ActionItem																														*
		//*-----------------------------------------------------------------------*
		private FileActionItem mActionItem = new FileActionItem();
		/// <summary>
		/// Get/Set the file action item associated with this session.
		/// </summary>
		public FileActionItem ActionItem
		{
			get { return mActionItem; }
			set { mActionItem = value; }
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	Run																																		*
		//*-----------------------------------------------------------------------*
		/// <summary>
		/// Run the configured application.
		/// </summary>
		public async Task Run()
		{
			await mActionItem.Run();
		}
		//*-----------------------------------------------------------------------*

		//*-----------------------------------------------------------------------*
		//*	WaitAfterEnd																													*
		//*-----------------------------------------------------------------------*
		private bool mWaitAfterEnd = false;
		/// <summary>
		/// Get/Set a value indicating whether to wait for user keypress after
		/// processing has completed.
		/// </summary>
		public bool WaitAfterEnd
		{
			get { return mWaitAfterEnd; }
			set { mWaitAfterEnd = value; }
		}
		//*-----------------------------------------------------------------------*

	}
	//*-------------------------------------------------------------------------*

}