using System;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows.Forms;

namespace NC_Office_Wrapper
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //only do something, if at leat 1 argument is given   
            if (args.Length > 0)
            {
                //where to find the path to the MS Office tools
                const string regWinword = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe";
                const string regExcel = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe";
                const string regVisio = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\visio.exe";
                string wordExec = null;     // used to store path to winword.exe
                string excelExec = null;    // used to store path to excel.exe
                string visioExec = null;    // used to store path to visio.exe

                //the lock-file we are looking for
                try
                {
                    var tempTuple = GetTempfileName(args[0]);
                    string filetype = tempTuple.Item1;
                    string lockFile = tempTuple.Item2;
                    string path = Path.GetDirectoryName(args[0]);
                    string filename = "\"" + args[0] + "\""; //we need to put " around the filename so everything is escape propperly

                    if (File.Exists(Path.Combine(path, lockFile)))
                    //if lock-file exists, show "file already in use" message to user
                    {
                        ShowMessage("File is currently locked.", "info");
                        Application.Exit();
                    }
                    else
                    // if there is no lock-file, chack filetype and open file with corresponding programm.
                    // if programm is not found, it's propably not installed. Show message about that.
                    // if filetype is unknowd, show message
                    {
                        switch (filetype)
                        {
                            case "word":
                                try
                                {
                                    wordExec = Registry.GetValue(regWinword, null, false).ToString();
                                    Process.Start(wordExec, filename);
                                }
                                catch
                                {

                                    ShowMessage("Microsoft Word was not found.");
                                }
                                break;
                            case "excel":
                                try
                                {
                                    excelExec = Registry.GetValue(regExcel, null, false).ToString();
                                    Process.Start(excelExec, filename);
                                }
                                catch
                                {

                                    ShowMessage("Microsoft Excel was not found.");
                                }
                                break;
                            case "visio":
                                try
                                {
                                    visioExec = Registry.GetValue(regVisio, null, false).ToString();
                                    Process.Start(visioExec, filename);
                                }
                                catch
                                {
                                    ShowMessage("Microsoft Visio was not found.");
                                }
                                break;
                            case "unknown":
                                ShowMessage("Filetype not supported");
                                break;


                        }
                    }
                }
                catch
                {
                    ShowMessage();
                }


            }
            else
            {
                // if less than one argument at wrapper-start, show a message
                ShowMessage("Please provide an Office-File.", "info");
            }
        }
        //Returns a Tuple with type of file and the name of the lock-file depending on the extention
        static Tuple<string, string> GetTempfileName(string file)
        {
            string filename_woe = Path.GetFileNameWithoutExtension(file);
            string extention = Path.GetExtension(file);
            string type = null;

            switch (extention.ToLower())
            {
                case ".doc":
                case ".docx":
                case ".docm":
                    type = "word";
                    filename_woe = CreateTempWord(filename_woe, extention);
                    break;
                case ".vsd":
                    type = "visio";
                    filename_woe = CreateTempVisio(filename_woe, extention);
                    break;
                case ".xlsx":
                case ".xlsm":
                    type = "excel";
                    filename_woe = CreateTempExcel(filename_woe, extention);
                    break;
                default:
                    type = "unknown";
                    filename_woe = "";
                    break;
            }

            return Tuple.Create(type, filename_woe);
        }

        //FOR .DOC AND .DOCX FILES
        //Returns the filename of the hidden .doc or docx lock-file. 
        //Takes in the original filename without extention and the extention
        static string CreateTempWord(string filenameoe, string extention)
        {
            //since this is a .doc or .docx file, we have to check how many chars we have to 
            //strip off the filename
            int toStrip = CalculateToStrip(filenameoe);

            filenameoe = filenameoe.Substring(toStrip);

            string tempfileName = "~$" + filenameoe + extention;

            return tempfileName;
        }

        //FOR .VSD FILES
        static string CreateTempVisio(string filenameoe, string extention)
        {
            //since this is a .vsd file, we have to check how many chars we have to 
            //strip off the filename
            int toStrip = CalculateToStrip(filenameoe);

            filenameoe = filenameoe.Substring(toStrip);

            //.vsd file's lockfiles get the extention ".~vsd"
            extention = ".~" + extention.Substring(1);

            string tempfileName = "~$$" + filenameoe + extention;

            return tempfileName;
        }

        //FOR .XLSX FILES (.xls files are stored in %TEMP% and are not usable)
        static string CreateTempExcel(string filenameoe, string extention)
        {
            string tempfileName = "~$" + filenameoe + extention;

            return tempfileName;
        }

        //Calculates the amount of characters to strip off the filename
        //Only applies to .doc, .docx and .vsd files
        static int CalculateToStrip(string filenameoe)
        {
            int toStrip;
            int length = filenameoe.Length;

            //Filenames with a length of 6 or less chars, don't get striped off first chars
            //Filenames with a length of 7 chars, get striped off 1 char
            //Filenames with a length of 8 or more chars, get striped off 2 char
            if (length < 7) { toStrip = 0; }
            else if (length < 8) { toStrip = 1; }
            else { toStrip = 2; }

            return toStrip;
        }

        //Shows a MessageBox 
        //Takes a message and a message level
        //Default is an error message
        static void ShowMessage(string message = "An error has occurred!", string level = "error")
        {
            MessageBoxIcon icon = default(MessageBoxIcon);

            switch (level)
            {
                case "error":
                    icon = MessageBoxIcon.Error;
                    break;
                case "info":
                    icon = MessageBoxIcon.Asterisk;
                    break;
            }

            MessageBox.Show(message,
                            "NC-Office-Wrapper",
                            MessageBoxButtons.OK,
                            icon,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.ServiceNotification);
            Application.Exit();
        }
    }
}
