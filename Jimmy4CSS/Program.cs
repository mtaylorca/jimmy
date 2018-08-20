using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Data;
using System.Xml.Linq;
using System.Runtime.InteropServices;

namespace Jimmy4CSS
{
    class Program
    {
        static readonly string DefaultPathToLaunchpad = @"C:\Program Files (x86)\Curtis Instruments\Integrated Toolkit\Launchpad.exe";
        static readonly string DefaultPathToDeviceProfiler = @"C:\Program Files (x86)\Curtis Instruments\Device Profiler\DeviceProfiler.exe";

        static readonly List<string> DefaultIgnoreFilesList = new List<string>(new string[] { "live.cmnu", "systemfullmenu.cmnu", "factorymenu.cmnu" });

        static readonly string XmlGenericSettingsFilename = "settings.xml";
        static readonly string XmlProjectSettingsString = "ProjectSettings";
        static readonly string XmlPathToLaunchpadString = "PathToLaunchpad";
        static readonly string XmlPathToDeviceProfilerString = "PathToDeviceProfiler";
        static readonly string XmlOutputDirectoryString = "OutputDirectory";
        static readonly string XmlIgnoreFileString = "IgnoreFile";

        static private void CreateXML(string i_XmlFilePath, string i_OutputDirectory, string i_PathToLaunchpad, string i_PathToDeviceProfiler, List<string> i_IgnoreFilesList)
        {
            Console.WriteLine("Creating config: " + i_XmlFilePath);

            using (XmlTextWriter writer = new XmlTextWriter(i_XmlFilePath, System.Text.Encoding.UTF8))
            {
                writer.WriteStartDocument(true);
                writer.Formatting = Formatting.Indented;
                writer.Indentation = 2;
                writer.WriteStartElement(XmlProjectSettingsString);

                writer.WriteStartElement(XmlPathToLaunchpadString);
                writer.WriteString(i_PathToLaunchpad);
                writer.WriteEndElement();

                writer.WriteStartElement(XmlPathToDeviceProfilerString);
                writer.WriteString(i_PathToDeviceProfiler);
                writer.WriteEndElement();

                writer.WriteStartElement(XmlOutputDirectoryString);
                writer.WriteString(i_OutputDirectory);
                writer.WriteEndElement();

                foreach (string ignoreFile in i_IgnoreFilesList)
                {
                    writer.WriteStartElement(XmlIgnoreFileString);
                    writer.WriteString(ignoreFile);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Close();
            }
        }

        static private bool ReadXML(string i_XmlFilePath, out string o_OutputDirectory, out string o_PathToLaunchpad, out string o_PathToDeviceProfiler, out List<string> o_IgnoreFilesList)
        {
            o_OutputDirectory = string.Empty;
            o_PathToLaunchpad = string.Empty;
            o_PathToDeviceProfiler = string.Empty;
            o_IgnoreFilesList = new List<string>();

            if (File.Exists(i_XmlFilePath) == false)
            {
                return false;
            }

            try
            {
                XElement xmlMap = XElement.Load(i_XmlFilePath);

                XName elementName = XName.Get(XmlOutputDirectoryString);
                o_OutputDirectory = xmlMap.Elements(elementName).First().Value;

                elementName = XName.Get(XmlPathToLaunchpadString);
                o_PathToLaunchpad = xmlMap.Elements(elementName).First().Value;

                elementName = XName.Get(XmlPathToDeviceProfilerString);
                o_PathToDeviceProfiler = xmlMap.Elements(elementName).First().Value;

                elementName = XName.Get(XmlIgnoreFileString);
                foreach (var ignoreFile in xmlMap.Elements(elementName))
                {
                    //Keep file names in lower case.
                    o_IgnoreFilesList.Add(ignoreFile.Value.ToLower());
                }
            }
            catch
            {
                //Failed to read from the XML.
                return false;
            }

            if (string.IsNullOrEmpty(o_OutputDirectory) || string.IsNullOrEmpty(o_PathToLaunchpad))
            {
                //These settings are required, so we will need to create them.
                return false;
            }
            else
            {
                return true;
            }
        }

        static void Main(string[] args)
        {
            //Disable quick edit mode, this where clicking on the console app will pause code execution...
            DisableQuickEdit();

            if (args.Length != 1)
            {
                Console.WriteLine("You must pass in the full path to a .CPRJ file.");
                Console.ReadKey();
                return;
            }

            string projectFilePath = args[0];

            string projectDirectory = Path.GetDirectoryName(projectFilePath);
            string fileName = Path.GetFileNameWithoutExtension(projectFilePath);            

            string outputDirectory;
            string pathToLaunchpad;
            string pathToDeviceProfiler;
            List<string> ignoreFilesList;

            //Look for the generic XML first
            string xmlFilePath = Path.Combine(projectDirectory, XmlGenericSettingsFilename);

            bool readOK = ReadXML(xmlFilePath, out outputDirectory, out pathToLaunchpad, out pathToDeviceProfiler, out ignoreFilesList);

            if (readOK == false)
            {
                //Look for the project name as the XML.
                xmlFilePath = Path.Combine(projectDirectory, fileName + ".xml");
                readOK = ReadXML(xmlFilePath, out outputDirectory, out pathToLaunchpad, out pathToDeviceProfiler, out ignoreFilesList);
            }

            if (readOK == false)
            {
                //Create a default XML.
                outputDirectory = fileName;
                pathToLaunchpad = DefaultPathToLaunchpad;
                pathToDeviceProfiler = DefaultPathToDeviceProfiler;
                ignoreFilesList = DefaultIgnoreFilesList;

                CreateXML(xmlFilePath, outputDirectory, pathToLaunchpad, pathToDeviceProfiler, ignoreFilesList);
            }

            //The output directory is always in the project directory.
            outputDirectory = Path.Combine(projectDirectory, outputDirectory);

            if (File.Exists(pathToLaunchpad) == false)
            {
                Console.WriteLine("Cannot find '" + pathToLaunchpad + "'.");
                Console.WriteLine("Please check your XML settings.");
                Console.ReadKey();
                return;
            }

            if (File.Exists(pathToDeviceProfiler) == false)
            {
                Console.WriteLine("Cannot find '" + pathToDeviceProfiler + "'.");
                Console.WriteLine("Please check your XML settings.");
                Console.ReadKey();
                return;
            }

            Process cssProcess = null;

            if (ProgramIsRunning(pathToLaunchpad))
            {
                Console.WriteLine("***Launchpad is already running. Please close it first as the wrong project may be open.***");
                Console.WriteLine("***If you are sure the correct project is open, press SPACE.***");

                if (Console.ReadKey().Key != ConsoleKey.Spacebar)
                {
                    return;
                }
            }
            else
            {
                Console.WriteLine("Opening: " + projectFilePath);

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = pathToLaunchpad;
                startInfo.Arguments = "\"" + projectFilePath + "\"";
                startInfo.UseShellExecute = false;
                cssProcess = Process.Start(startInfo);
            }

            //Create the output directory if it doesnt already exist.
            Console.WriteLine("Output directory: " + outputDirectory);
            Directory.CreateDirectory(outputDirectory);
            
            string directoryToWatch = Path.Combine(Path.GetTempPath(), "C2IT");

            CSSFileWatcher fileWatcher = new CSSFileWatcher(directoryToWatch, outputDirectory, pathToDeviceProfiler, ignoreFilesList, cssProcess);
            fileWatcher.RunMainCode();
        }


        static private bool ProgramIsRunning(string FullPath)
        {
            string FilePath = Path.GetDirectoryName(FullPath);
            string FileName = Path.GetFileNameWithoutExtension(FullPath).ToLower();
            bool isRunning = false;

            try
            {
                Process[] pList = Process.GetProcessesByName(FileName);

                foreach (Process p in pList)
                {
                    if (p.MainModule.FileName.StartsWith(FilePath, StringComparison.InvariantCultureIgnoreCase))
                    {
                        isRunning = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return isRunning;
        }

        /**********************************************************
        Code to mess with the command prompt down here...
        **********************************************************/

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetConsoleMode(
            IntPtr hConsoleHandle,
            out int lpMode);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool SetConsoleMode(
            IntPtr hConsoleHandle,
            int ioMode);

        public const int STD_INPUT_HANDLE = -10;

        [DllImport("Kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetStdHandle(int nStdHandle);

        /// <summary>
        /// This flag enables the user to use the mouse to select and edit text. To enable
        /// this option, you must also set the ExtendedFlags flag.
        /// </summary>
        const int QuickEditMode = 64;

        // ExtendedFlags must be combined with
        // InsertMode and QuickEditMode when setting
        /// <summary>
        /// ExtendedFlags must be enabled in order to enable InsertMode or QuickEditMode.
        /// </summary>
        const int ExtendedFlags = 128;

        private static void DisableQuickEdit()
        {
            IntPtr conHandle = GetStdHandle(STD_INPUT_HANDLE);
            int mode;

            if (!GetConsoleMode(conHandle, out mode))
            {
                // error getting the console mode. Exit.
                return;
            }

            mode = mode & ~(QuickEditMode | ExtendedFlags);

            if (!SetConsoleMode(conHandle, mode))
            {
                // error setting console mode.
            }
        }

        private static void EnableQuickEdit()
        {
            IntPtr conHandle = GetStdHandle(STD_INPUT_HANDLE);
            int mode;

            if (!GetConsoleMode(conHandle, out mode))
            {
                // error getting the console mode. Exit.
                return;
            }

            mode = mode | (QuickEditMode | ExtendedFlags);

            if (!SetConsoleMode(conHandle, mode))
            {
                // error setting console mode.
            }
        }
    }

    class CSSFileWatcher
    {
        private readonly string DirectoryToWatch;
        private readonly string OutputDirectory;
        private readonly string PathToDeviceProfiler;
        private readonly List<string> IgnoreFilesList;
        private readonly Process CssProcess;

        /// <summary>
        /// Keep a dictionary for filenames with the path to their latest version.
        /// </summary>
        private Dictionary<string, string> LatestFiles = new Dictionary<string, string>();        

        public CSSFileWatcher(string i_DirectoryToWatch, string i_OutputDirectory, string i_PathToDeviceProfiler, List<string> i_IgnoreFilesList, Process i_CssProcess)
        {
            this.DirectoryToWatch = i_DirectoryToWatch;
            this.OutputDirectory = i_OutputDirectory;
            this.PathToDeviceProfiler = i_PathToDeviceProfiler;
            this.IgnoreFilesList = i_IgnoreFilesList;
            this.CssProcess = i_CssProcess;
        }

        //private Application ExcelApp = new Application();

            /*
        private void FormatM4(string i_Path)
        {
            if (File.Exists(i_Path) == false)
            {
                Console.Write("File not found: " + i_Path);
                return;
            }

            if (this.ExcelApp == null)
            {
                //Only create once.
                this.ExcelApp = new Application();
            }
            
            //ExcelApp.Visible = true;
            this.ExcelApp.DisplayAlerts = false;

            try
            {

                Workbook workbook = ExcelApp.Workbooks.Open(i_Path,
                                                            Type.Missing,
                                                            Type.Missing,
                                                            XlFileFormat.xlCSV,   // Format
                                                            Type.Missing,
                                                            Type.Missing,
                                                            Type.Missing,
                                                            Type.Missing,
                                                            ',',          // Delimiter
                                                            Type.Missing,
                                                            Type.Missing,
                                                            Type.Missing,
                                                            Type.Missing,
                                                            Type.Missing,
                                                            Type.Missing);

                Worksheet worksheet = (Worksheet)workbook.ActiveSheet;

                //Delete the index row
                int columnCount = worksheet.UsedRange.Columns.Count;
                int rowCount = worksheet.UsedRange.Rows.Count;
                bool indexRowFound = false;

                List<string> columnNames = new List<string>();

                for (int c = 1; c < columnCount; c++)
                {
                    string columnName = worksheet.Cells[1, c].Value.ToString();

                    if (columnName == "Index")
                    {
                        //Found the index column, delete it.
                        //Console.WriteLine("Deleting Index Column");
                        Range range = worksheet.Columns[c];
                        range.Delete();
                        indexRowFound = true;
                        break;
                    }
                }

                if (indexRowFound == false)
                {
                    //The index row was not found, this might be the wrong file type. Or it was already processed.                
                    Console.Write("Error with " + Path.GetFileName(i_Path) + "...");

                    workbook.Close(0);
                    ExcelApp.Quit();
                    return;
                }

                Range oRng = worksheet.Range["A1"];

                oRng.EntireColumn.Insert(XlInsertShiftDirection.xlShiftToRight, XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                oRng = worksheet.Range["A1"];

                oRng.Value2 = "AddressOrName";

                //Console.WriteLine("Applying Formula");
                double progressSteps = rowCount / 10;
                int progressCounter = 0;
                Console.Write("10...");

                for (int currentRow = 2; currentRow < rowCount + 1; currentRow++)
                {
                    //Skip first row
                    //Console.WriteLine(r + "/" + rowCount);
                    if (currentRow > progressSteps * progressCounter && progressCounter < 10)
                    {
                        Console.Write((10 - ++progressCounter) + "...");
                    }

                    //Ex: =IF(N2="(null)",A2,N2)
                    string formula = "=IF(N" + currentRow + "=\"(null)\",B" + currentRow + ",N" + currentRow + ")";

                    worksheet.Cells[currentRow, 1].Formula = formula;
                }

                workbook.SaveAs(i_Path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook.Close(0);
                ExcelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

                this.ExcelApp = null;

                //Ensure excel is nuked.
                var process = System.Diagnostics.Process.GetProcessesByName("Excel");
                foreach (var p in process)
                {
                    if (!string.IsNullOrEmpty(p.ProcessName))
                    {
                        try
                        {
                            p.Kill();
                        }
                        catch { }
                    }
                }
            }
        }
        */

        public void RunMainCode()
        {            
            while(Directory.Exists(this.DirectoryToWatch) == false)
            {
                //Waiting for directory to be created.
            }

            try
            {
                //Create a new FileSystemWatcher and set its properties.
                FileSystemWatcher watcher = new FileSystemWatcher();
                watcher.Path = this.DirectoryToWatch;

                //Watch both files and subdirectories.
                watcher.IncludeSubdirectories = true;

                //Watch for all changes specified in the NotifyFilters enumeration.
                watcher.NotifyFilter = NotifyFilters.Attributes |
                NotifyFilters.CreationTime |
                NotifyFilters.DirectoryName |
                NotifyFilters.FileName |
                NotifyFilters.LastAccess |
                NotifyFilters.LastWrite |
                NotifyFilters.Security |
                NotifyFilters.Size;

                //Watch all files.
                watcher.Filter = "*.*";

                //Add event handlers.
                watcher.Changed += new FileSystemEventHandler(OnChanged);

                //Start monitoring.
                watcher.EnableRaisingEvents = true;

                Console.WriteLine("Jimmy has started.");
                Console.Write(this.Jimmy);
                Console.WriteLine();

                while (this.ForceShutdown == false)
                {
                    //Make an infinite loop till 'ESC' is pressed. Or if CSS has been closed.
                    if (this.CssProcess != null && this.CssProcess.HasExited)
                    {
                        this.ForceShutdown = true;
                    }

                    if (Console.KeyAvailable)
                    {
                        //Only call readkey if a key is available, otherwise this is a blocking call.

                        if (Console.ReadKey(true).Key == ConsoleKey.Escape)
                        {
                            this.ForceShutdown = true;
                        }
                    }
                }

                Console.WriteLine("Shutting down.");
                while (ProcessingChanges)
                {
                    //Wait until processing has completed.
                }
            }
            catch (IOException e)
            {
                Console.WriteLine("A Exception Occurred :" + e);

                Console.WriteLine("Press any key to quit.");
                Console.ReadKey();
            }

            catch (Exception oe)
            {
                Console.WriteLine("An Exception Occurred :" + oe);

                Console.WriteLine("Press any key to quit.");
                Console.ReadKey();
            }
        }

        private bool ForceShutdown = false;
        private bool ProcessingChanges = false;

        // Define the event handlers.
        private void OnChanged(object source, FileSystemEventArgs e)
        {
            ProcessingChanges = true;

            FileInfo newFile = new FileInfo(e.FullPath);

            if (newFile.Extension.ToLower() == ".cmnu" || newFile.Extension.ToLower() == ".vcl")
            {
                //Good file
                //Console.WriteLine(newFile.Name);
            }
            else
            {
                //Ignore other file types
                ProcessingChanges = false;
                return;
            }

            if (this.IgnoreFilesList.Contains(newFile.Name.ToLower()))
            {
                //Ignore this file.
                ProcessingChanges = false;
                return;
            }

            long fileLength = newFile.Length;
            DateTime fileWriteTime = newFile.LastWriteTime;
            
            if (this.LatestFiles.Keys.Contains(newFile.Name))
            {
                //File already exists in the dictionary.

                FileInfo existingFile = new FileInfo(this.LatestFiles[newFile.Name]);

                if (newFile.LastWriteTime > existingFile.LastWriteTime)
                {
                    //Newer file.
                    this.LatestFiles[newFile.Name] = e.FullPath;
                }
                else
                {
                    //This file is old. Skip it.
                    ProcessingChanges = false;
                    return;
                }
            }
            else
            {
                //Add a new one
                this.LatestFiles.Add(newFile.Name, e.FullPath);                
            }

            //Process the file
            Console.WriteLine("{0}: {1} - {2} bytes - {3} ", e.ChangeType, Path.GetFileName(e.FullPath), fileLength, fileWriteTime.ToLongTimeString());

            while (FileIsLocked(e.FullPath, FileAccess.Read))
            {
                //Sit here waiting!
            }

            string newFilePath = Path.Combine(this.OutputDirectory, newFile.Name);
            newFile.CopyTo(newFilePath, true);

            if (newFile.Extension.ToLower() == ".cmnu")
            {
                Console.Write("Processing...");
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = this.PathToDeviceProfiler;
                startInfo.Arguments = "-I\"" + newFilePath + "\" -O\"" + newFilePath + ".M4.CSV\"" + " -O\"" + newFilePath + ".P4.CSV\"";
                startInfo.CreateNoWindow = true;
                startInfo.UseShellExecute = false;
                Process.Start(startInfo).WaitForExit();

                /*
                //Now diddle the M4 file
                string m4FilePath = newFilePath + ".M4.CSV";
                FormatM4(m4FilePath);
                */

                Console.WriteLine("Done.");
            }

            ProcessingChanges = false;
        }

        // Return true if the file is locked for the indicated access.
        private bool FileIsLocked(string filename, FileAccess file_access)
        {
            // Try to open the file with the indicated access.
            try
            {
                FileStream fs = new FileStream(filename, FileMode.Open, file_access);
                fs.Close();
                fs.Dispose();

                return false;
            }
            catch (IOException)
            {
                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private readonly string Jimmy = @"                             
             ▄▄▄▄▄▄▄         
         ▄▀▀▀░░░░░░░▀▄       
       ▄▀░░░░░░░░░░░░▀▄      
      ▄▀░░░░░░░░░░▄▀▀▄▀▄     
    ▄▀░░░░░░░░░░▄▀░░██▄▀▄    
   ▄▀░░▄▀▀▀▄░░░░█░░░▀▀░█▀▄   
   █░░█▄▄░░░█░░░▀▄░░░░░▐░█   
  ▐▌░░█▀▀░░▄▀░░░░░▀▄▄▄▄▀░░█  
  ▐▌░░█░░░▄▀░░░░░░░░░░░░░░█  
  ▐▌░░░▀▀▀░░░░░░░░░░░░░░░░▐▌ 
  ▐▌░░░░░░░░░░░░░░░▄░░░░░░▐▌ 
  ▐▌░░░░░░░░░▄░░░░░█░░░░░░▐▌ 
   █░░░░░░░░░▀█▄░░▄█░░░░░░▐▌ 
   ▐▌░░░░░░░░░░▀▀▀▀░░░░░░░▐▌ 
    █░░░░░░░░░░░░░░░░░░░░░█  
    ▐▌▀▄░░░░░░░░░░░░░░░░░▐▌  
     █  ▀░░░░░░░░░░░░░░░░▀   
                             ";

    }
}
