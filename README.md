# FileTools Console Application

This is a .NET console application that provides advanced batch-style operations on files.

The current version is Windows-only, due to its use of the **System.Drawing** namespace. However, this will soon be written out of the application so I can provide full cross-platform functionality. Please stay tuned if you are interested in using this application on Linux and macOS.

<p>&nbsp;</p>

## 🆕 LibreOffice Calc to Generic JSON

The action ConvertCalcToJson has been added to convert a LibreOffice Calc ODS file to a general, flat JSON format. If all of the non-blank values in a column are numeric, the output property will be treated as numeric. Similarly, if all of the non-blank values in a column are boolean, the output property will be treated as boolean.

<p>&nbsp;</p>

## 🆕 LibreOffice Calc Blender Timeline to Blender Keyframe JSON Script

The action **ConvertCalcToBlenderKeyframes** has been added to convert a LibreOffice Calc ODS file using a Blender timeline animation layout to a Blender keyframe JSON file compatible with the ImportKeyframes.py script on the GitHub project [danielanywhere/BlenderAnimationImport](https://github.com/danielanywhere/BlenderAnimationImport).

<p>&nbsp;</p>

## Syntax

Following is the current application syntax.

```plaintext
Command-line tool suite for working with files.

Syntax:
FileTools.exe /action:{ActionName}
    [/configfile:{Filename}]
    [/input:{Filename|FolderName}]
    [/output:{Filename|FolderName}]
    [/pattern:{Pattern}]
    [/infolder:{FolderName}] [/outfolder:{FolderName}]
    [/infile:{Filename}] [/outfile:{Filename}]
    [/option:{OptionName[,OptionValue]}]
    [/range:{Start},{End}] [/base:{Base}]
    [/digits:{Digits}] [/count:{Count}] [/datetime:{DateTime}]
    [/sheetname:{Name}]
    [/text:{Text}]
    [/prefix] [/suffix]
    [/properties:{NameValueArray}]
    [/workingpath:{FolderName}]
    [/recurse]
    [/wait]
    [/?]

 * -    Actions marked with an asterisk are only supported within a
        configuration file.

    /action     -   Describes the action to be made. Following are the
                    recognized actions, with associated parameter names.
                    Starred names ('*') are only available in a batch.
                    Batch - Perform a batch of file operations from a
                        single JSON configuration file.
                        /infile
                    BuildPathProperty * - Build an absolute path in a user
                        property from a partial path template.
                    ClearInputFiles * - Clear the input files collection at
                        this level.
                    ConvertCalcToBlenderKeyframes - Convert a LibreOffice Calc
                        ODS file to a Blender keyframe JSON file compatible
                        with the ImportKeyframes.py script on the GitHub
                        project danielanywhere/BlenderAnimationImport.
                        /infile, /sheetname, /outfile
                        [
                            "/properties:[{ 'Name':'ExcludeObjects','Value':
                            'ObjectName1,ObjectName2,etc' }]"
                        ]
                    ConvertCalcToJson - Convert a LibreOffice Calc ODS file
                        to JSON.
                        /infile, /sheetname, /outfile
                    ConvertFromB64 - Convert a file from base-64 to binary.
                        /infile, /outfile
                    ConvertToB64 - Convert a file from binary to base-64.
                        Available options: DataUrl,{true|false}
                        /infile, /outfile
                    CopyFirstNumeric - Copy the first numeric file in the
                        source folder to the target.
                        /input, /output
                    CopyIncrementReverse - Copy a range of files in the source
                        folder to a higher increment, in reverse order.
                        /infolder, /range, /outfolder, /base
                    CopyLastNumeric - Copy the last numbered file in the
                        source folder to the target.
                        /infolder, /output
                    CopyLastNumericExtend - Copy the last numbered file in the
                        source folder to an extended number of frames.
                        /infolder, /outfolder, /base
                    CopyMergeSubs - Copy all files from subfolders to the
                        output folder, using merged sequential naming.
                        /infolder, /outfolder, /base
                    CopyNumericToRange - Copy the specfied numeric frame to the
                        specified range.
                        /infile, /outfolder, /range, [/digits]
                    CopyRange - Copy a range of files in the source folder to
                        a new destination.
                        /infolder, /range, [/outfolder], [/base], [/digits]
                    CropImage * - Crop the working image to the provided Left,
                        Top, Width, and Height values.
                    CropImageToRectangleInfoName * - Crop the working image to
                        the rectangle specified in the item of the rectangle
                        list corresponding to the provided Name user property.
                    DateCopy - Copy files or folders from one location to
                        another, adding a date prefix or suffix to each file
                        at the destination location.
                        /input, [/prefix | /suffix], /outfolder
                    DelDirectoryPattern - Delete every directory matching the
                        specified pattern in the input folder.
                        /infolder
                    DeleteFile - Delete the file specified in the Filename
                        user property if it exists.
                    DelEveryX - Delete every Xth file in the selected files
                        list.
                        /input, /count
                    DirReformat - Reformat the results of a DIR output
                        file to CSV.
                        /infile, /outfile
                    DirToTsv - Scan a directory and save the findings as a
                        tab-separated values (TSV) file.
                        /infolder, /outfile
                    DrawImage * - Draw the image specified by ImageName onto
                        the working image at the location specified by user
                        properties Left and Top.
                    FileOpenImage * - Open the image file specified in the
                        current input file. Name it in the local images
                        collection with the name specified in the user
                        property ImageName.
                    FileOverlayImage * - Open each image from the range and
                        place the image specified in InputFilename at the
                        options specified by Left, Top, Width, and Height.
                        {/infolder, /range|/input}
                    FileSaveImage * - Save the working image to the currently
                        specified OutputFile.
                    FindFiles - Find the files matching the provided pattern.
                        Properties: 'Find'
                        /input /recurse
                    ForEachFile * - Run the Actions collection of the action
                        through all of the files currently loaded in the
                        InputFiles collection, setting the CurrentFile
                        property for each pass.
                    FormatDirFile - Format the text output of a DOS style
                        command-line DIR /s command to create a tab-separated
                        values (TSV) file.
                        /infile, /outfile
                    ImagesClear * - Clear all images from the Images
                        collection.
                    ImageSetCommonBoundary - Create common bounding boxes for
                        the image sets specified. If text is specified, it is
                        used as the name of the output structure.
                        Available options are Square and Grow,{Amount}
                        /input, [/text], [/option]
                    ListAllFilesOnDate - List all files created or modified
                        within the specified date range.
                        /infolder, /range
                    LoadRectangleInfoList * - Load the Rectangle info list
                        from an external JSON file.
                    MoveFiles - Move the specified files from one folder to
                        another.
                        /input, /outfolder
                    PrefixFilenames - Prefix matching filenames with a
                        text pattern.
                        /input, /text
                    RenameFiles - Rename matching files with regular
                        expression find and replace.
                        Properties: 'Find', 'Replace'
                        /input, /output, /recurse
                    RenumberFiles - Renumber all of the matching files to be
                        contiguous.
                        /infolder, [/outfolder], [/range], [/base], [/digits]
                    RepeatInsertClip - Repeat a group of one or more files
                        sequentially. If base is omitted, the starting target
                        is the next number after the end of the source clip.
                        /infolder, [/outfolder], /range, [/base]
                    ReplaceGreenscreen - Remove background from one or more
                        image files, placing the results in a separate
                        folder. Uses Microsoft Cognitive Services and requires
                        the properties MSCognitiveServicesKey and
                        MSCognitiveServicesEndpoint.
                        /input, /outfolder
                    RunSequence * - Run the sequence specified in the
                        'SequenceName' user property.
                    SetFileDate - Set the file date and time (TOUCH) on the
                        specified files.
                        /input, /datetime
                    SetWorkingImage * - Set the current working image to the
                        one with the local name found in the user property
                        ImageName.
                    SizeImage * - Scale the image to a new size, as specified
                        in user properties Width and Height.
                    StitchFilePatternToMp4 - Stitch files found in the
                        FilePattern property to the output filename using
                        FFMPEG.
                    SuffixFilenames - Suffix matching filenames with a
                        text pattern.
                        /input, /text
    /base       -   The base number or filename pattern of the source or
                    target files, depending upon the current action.
    /count      -   Count of interations to take on the current action.
    /datetime   -   The date and time to apply to the action.
    /digits     -   Count of digits to output.
    /infile     -   Input path and filename only.
    /infolder   -   Input path and folder name only.
    /input      -   An input pattern that allows for filenames or foldernames
                    with or without wildcards. This parameter can be specified
                    multiple times on the command line with different values
                    to load multiple input files.
                    Since input is the most generic type of specification, if
                    an input parameter is allowed on an action, the infile or
                    infolder parameters can be specified once in its place if
                    the input parameter itself is left blank.
    /option     -   Specifies a single option name and optional matching
                        value. Multiple options can be specified per command.
                        Options are specific to the action context in which
                        they are specified.
    /outfile    -   Output path and filename only.
    /outfolder  -   Output path and folder name only.
    /output     -   An output pattern that allows for filenames or foldernames
                    with or without wildcards. This parameter can be specified
                    muliple times on the command line with different values
                    to write to multiple output files.
    /pattern    -   Match the supplied regular expression pattern for files,
                    folders, or other appropriate strings.
    /prefix     -   A flag indicating that the prefix function is active
                    for the current action.
    /range      -   A range of values in the format suitable for the context
                    of the action. This parameter will most likely be
                    composed of two numbers, dates, or strings.
    /recurse    -   When specified, the action will recurse inward to all
                    subdirectories from the current working folder.
                    This flag is currently only recognized for the RenameFiles
                    action.
    /suffix     -   A flag indicating that the suffix function is active
                    for the current action.
    /workingpath-   Set the working path and foldername. When this parameter
                    is specified, all of the other parameters can use relative
                    naming.

{ActionName}    -   Name of the action to execute.
{Base}          -   The base number to use, typically as a target value.
{DateTime}      -   Date and / or Time, in one of the formats of
                    MM/DD/YYYY[ HH:MM[:SS[ FFF]]] or YYYYMMDD[.HHMM[SS[FFF]]].
{Digits}        -   Count of characters or digits.
{End}           -   End pattern of a range.
{Filename}      -   Fully qualified path and filename.
{Foldername}    -   Fully qualified path.
{OptionName}    -   Name of the option. Not case sensitive.
{OptionValue}   -   Value to apply to the option. Some options do not require
                        values.
{Pattern}       -   File or folder regular expression match pattern.
{Start}         -   Start pattern of a range.
{Text}          -   A non-regex text pattern to apply.

Options:
Following are the available options.

DataUrl         -   Indication of whether to create a Data URL from the
                    current data.
                    Value: {true|false}
                    Commands: ConvertToB64
Grow            -   Grow the image by the specified number of pixels.
                    Value: Total number of pixels to grow.
                    Commands: ImageSetCommonBoundary
Mute            -   Don't run this action.
                    Value: (none)
                    Commands: (Any)
SavePicture     -   Save a copy of the picture created by the process.
                    Value: Filename of the picture.
                    Commands: ImageSetCommonBoundary
Solo            -   Run only this action within the context of the current
                    parent batch.
                    Value: (none)
                    Commands: (Any)
Square          -   Square the image.
                    Value: (none)
                    Commands: ImageSetCommonBoundary

Example:
FileTools /action:CopyNumericToRange /input:C:\Temp\455.png /
    /range:5000,10000 /outfolder:C:\DestinationFolder

The above example copies frame 455 from the C: Temp folder to a 5 digit
sequence of files at C:\DestinationFolder ranging from 05000.png to
10000.png. Since the digits parameter isn't specified, the maximum width
of the target range is automatically selected.

Batch config file variables.

| Name | Type | Description |
|------|------|-------------|
| Action | ActionTypeEnum | The action to be made on the file. |
| Actions | List<ActionItem> | Collection of actions to run as a part of this
action. |
| Base | string | The base number or filename pattern of the source or target
files, depending upon the action. |
| Count | float | The count associated with the current action. |
| DateTimeValue | DateTime | the date and time associated with the current
action. |
| Digits | int | The number of digits associated with the current action. |
| InputFilename | string | The path and filename of the input file. |
| InputFolderName | string | The input path for the current operation. |
| Options | List<OptionItem> | Collection of operational options. |
| OutputFilename | string | The output path and filename for the operation. |
| OutputFolderName | string | The output path for the operation. |
| OutputName | string | Generic output path or filename. |
| OutputType | RenderFileType | type of rendering to be done on the affected
file. |
| Pattern | string | Active regular expression pattern. |
| Properties | List<PropertyItem> | Collection of properties for action. |
| Range | StartEndItem | The start and end values of the range. |
| Recurse | bool | Value indicating whether operations should recurse to sub-folders. |
| RectangleInfoList | RectInfoCollection | Rectangle list. |
| Sequences | List<SequenceItem> | Sequences for action. |
| Text | string | Text for the current action. |
| WorkingPath | string | Working path for current operation. |

```

<p>&nbsp;</p>

## Example Command-Lines

The following code block provides practical examples of how this utility is used.

```plaintext
Example command lines for FileTools.

filetools /wait /action:Batch "/configfile:C:\Develop\Projects\Simulations\Assembly\ElectricMotor\T80B34B\Data\FileToolsRotatingCropSizeLayout.json"

filetools /wait /action:Batch "/configfile:C:\Develop\AdvancedLogicLearningSystems\Scripts\StatingTheProblem01Rename01.json"

filetools /wait /action:ConvertToB64 "/properties:[{'Name':'DataUrl','Value':true}]" "/infile:C:\Develop\Projects\Simulations\Assembly\ElectricMotor\T80B34B\Images\PartImageAmatureShaft.png" "/outfile:C:\Develop\Projects\Simulations\Assembly\ElectricMotor\T80B34B\b64\PartImageAmatureShaft.b64"

filetools /wait /action:ConvertToB64 /option:DataUrl,true "/infile:C:\Develop\Projects\Simulations\Assembly\ElectricMotor\T80B34B\Images\PartImageAmatureShaft.png" "/outfile:C:\Develop\Projects\Simulations\Assembly\ElectricMotor\T80B34B\b64\PartImageAmatureShaft.b64"

filetools /wait /action:FileOverlayImage "/infolder:C:\Develop\Projects\ThumbMakerJS\Video1\RawFootage\Frames" /range:Frame-02084.png,Frame-02793.png "/option:MaskFilename,C:\Develop\Projects\ThumbMakerJS\Video1\RawFootage\MaskExWhite.png" /option:Left,556 /option:Top,302 /option:Width,147 /option:Height,27
filetools /wait /action:FileOverlayImage "/input:Output-Robert1\*.png" "/option:MaskFilename,C:\GreenScreenProject\RawFootage\RobertArtifact01.png" /option:Left,1024 /option:Top,672 /option:Width,240 /option:Height,32 "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"

filetools /wait /action:DirToTsv /option:Fields,Filename;Extension;RelativeDir /option:SeparateFilenameExtension,true /option:Recurse,true "/infolder:C:\Develop\GitHub\piranha.core\core" "/outfile:C:\Develop\Projects\AscendantDesignWeb\Web24\Docs\PiranhaTasks\CoreListing.txt"

filetools /wait /action:DelDirectoryPattern "/workingpath:C:\Develop\Develop\AscendantDesignWeb24\pCore" "/infolder:*\*\bin"
filetools /wait /action:DelDirectoryPattern "/workingpath:C:\Develop\Develop\AscendantDesignWeb24\pCore" "/infolder:*\*\obj"

filetools /wait /action:RenameFiles "/properties:[{ 'Name': 'Find', 'Value': '(?i:(?<name>[a-z0-9-_ ]+ slide (?<number>\\d{2})[^\\.]+\\.mp3))' },{ 'Name': 'Replace', 'Value': 'ELSlide${number}.mp3' }]" "/workingpath:C:\Develop\Training\Topic6_EffectiveLeadership\Video1\EffectiveLeadership_VO_Finalized"

filetools /wait /action:Batch "/configfile:..\Scripts\FileToolsRenameVoiceoverFiles.json" "/workingpath:C:\Develop\Training\Topic6_EffectiveLeadership\Video1\EffectiveLeadership_VO_Finalized"

filetools /wait /action:Batch "/configfile:Scripts\FileToolsNLECore.json" "/workingpath:C:\Develop\Training\5-Videos\Construction"

MSCognitiveServicesKey MSCognitiveServicesEndpoint
filetools /wait /action:RemoveBackground "/properties:[{'Name': 'MSCognitiveServicesKey','Value': 'abcd1234'},{'Name': 'MSCognitiveServicesEndpoint', 'Value':'visionservicename'}]" "/input:C:\Temp\TestImage.png" "/outfolder:C:\Temp\Movies"

filetools /wait /action:RenumberFiles "/infolder:C:\Temp\Movies\Frames\video1380591207_Intro" "/range:Frame-00001.png,Frame-01668.png" "/outfolder:Renumbered" "/base:Frame-00002.png" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"

filetools /wait /action:AlphaConditionalAdjust "/input:Transparent\*.png" "/outfolder:Alpha01" "/properties:[{'Name': 'Condition', 'Value': 'a<=195'},{'Name': 'Assignment', 'Value': '(a-32)/2'}]" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"
filetools /wait /action:AlphaConditionalAdjust "/input:Alpha01\*.png" "/outfolder:Alpha02" "/properties:[{'Name': 'Condition', 'Value': 'a=255 AND r<=97 AND g>=128 AND b<=97'},{'Name': 'Assignment', 'Value': '0'}]" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"
filetools /wait /action:AlphaConditionalAdjust "/input:Alpha02\*.png" "/outfolder:Alpha03" "/properties:[{'Name': 'Condition', 'Value': 'a=255 AND r<=110 AND g>=140 AND b<=105'},{'Name': 'Assignment', 'Value': '0'}]" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"
filetools /wait /action:AlphaConditionalAdjust "/input:Alpha03\*.png" "/outfolder:Alpha04" "/properties:[{'Name': 'Condition', 'Value': 'a=255 AND r<=100 AND g>=166 AND b<=121'},{'Name': 'Assignment', 'Value': '0'}]" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"

filetools /wait /action:AntiAliasTransparency "/input:Alpha04\*.png" "/outFolder:Alpha05" "/properties:[{'Name': 'Distance', 'Value': '5'}]" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"

filetools /wait /action:Batch "/configfile:C:\GreenScreenProject\Scripts\Experiment\FileToolsPrepareAlpha.json" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"

filetools /wait /action:Batch "/configfile:Scripts\FileToolsNLEAlphaProcessing.json" "/workingpath:C:\GreenScreenProject"

filetools /wait /action:Batch "/infile:C:\GreenScreenProject\Scripts\EditSheet.xlsx" "/configfile:C:\GreenScreenProject\Scripts\FileToolsNLEAlphaProcessing.json" "/workingpath:C:\Temp\Movies\Frames\video1380591207_Intro"

filetools /wait /action:Batch "/infile:C:\BumperVideoRevisions\Scripts\EditSheet.xlsx" "/configfile:C:\BumperVideoRevisions\Scripts\FileToolsNLEAlphaProcessing.json" "/workingpath:C:\Temp\Movies\Frames\ScreenshareVideo-2024081601"

filetools /wait /action:RenameFiles "/properties:[{ 'Name': 'Find', 'Value': '(?i:(?<number>\\d+)\\.png)' },{ 'Name': 'Replace', 'Value': 'Frame-${number}.png' }]" "/workingpath:C:\Temp\Movies\Frames\OpeningCNC"

filetools /wait /action:Batch "/configfile:C:\Develop\Projects\CNC\HobbyCNC\Scripts\ImageBackground.json" "/infile:C:\Develop\Projects\CNC\HobbyCNC\Scripts\EditSheet.xlsx" "/workingpath:C:\Develop\Projects\CNC\HobbyCNC"

filetools /wait /action:RenameFiles "/properties:[{ 'Name': 'Find', 'Value': '(?i:(?<number>\\d+)\\.png)' },{ 'Name': 'Replace', 'Value': 'Frame-${number}.png' }]" "/workingpath:C:\Temp\Movies\Frames\OpeningCNC"

filetools /wait /action:RenameFiles /recurse "/properties:[{ 'Name': 'Find', 'Value': 'Chat _ Bridget Manley _ Microsoft Teams \\(10_3_2024 11_22_27 AM\\)\\.html' },{ 'Name': 'Replace', 'Value': 'Chat-BridgetDaniel-20210827-20220323.html' }]" "/workingpath:C:\OneDrive\Ascendant\IT Department\TeamsBackup\Chats"

filetools /wait /action:FormatDirFile /infile:C:\Temp\DirAll.txt /outfile:C:\Temp\DirAll-Formatted.txt

filetools /wait /action:RenameFiles "/properties:[{ 'Name': 'Find', 'Value': '(?i:(?<name>image\\d+\\.[a-z]{3,4}))' },{ 'Name': 'Replace', 'Value': '%FILENAME%-${name}' }]" "/workingpath:%MEDIA%\media"

filetools /wait /action:FormatDirFile /infile:C:\Temp\DirAll.txt /outfile:C:\Temp\DirAll-Formatted.txt

filetools /wait /recurse /action:FindFiles /input:C:\Files\Dropbox\Develop\Active "/properties:[{'Name':'Find','Value': '(?i:^[^\\.]*svg[^\\.]*)'}]"

filetools /wait /action:ConvertCalcToJson /sheetname:TestSheet /workingpath:C:\Files\Dropbox /infile:MySpreadsheet.ods /outfile:ConvertedFromCalc.json

filetools /wait /action:ConvertCalcToBlenderKeyframes /sheetname:Timeline "/properties:[{ 'Name': 'ExcludeObjects', 'Value': 'LampPoint.001' }]" /workingpath:C:\Files\Dropbox /infile:Docs/BlenderTimeline.ods /outfile:Scripts/BlenderKeyframes.json

```
