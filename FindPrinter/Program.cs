using System;
using System.Diagnostics;
using System.Linq;
using System.Printing;
using System.Windows.Automation;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.CodeDom;
using System.Management;

namespace FindPrinters
{
    public class Program
    {

        [DllImport("user32.dll")] public static extern int FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")] public static extern int PostMessage(int hWnd, uint Msg, int wParam, int lParam);

        const uint WM_CLOSE = 0x0010;
        static void Main()
        {
            PrintServer printServer = new PrintServer();

            string selectedPrinter;

            //PrintQueueColllection
            PrintQueueCollection printQueues = printServer.GetPrintQueues();

            //convert to array
            PrintQueue[] printQueueArray = printQueues.ToArray();

            int optionToNumberConverter = 2;
            //do
            //{
            /*
            Console.WriteLine("----------Console---------------------");
            Console.WriteLine("Please Press 1 to List Printers: ");
            Console.WriteLine("Please Press 2 to Search For a Printer.");
            Console.WriteLine("Please Press 3 to Exit the Program.\n");
            Console.WriteLine("Type a Number and Hit Enter.");
            Console.WriteLine("--------------------------------------\n");

            Console.Write("Option: ");
            string UserInput = Console.ReadLine();

            Console.Clear();

            Console.WriteLine("User selected option:  " + UserInput);
            Console.WriteLine();
            */
            switch (optionToNumberConverter)
            {
                case 1:
                    Console.WriteLine("Standby.. Gathering Printer Information\n");

                    //displays printers below this writeline.
                    Console.WriteLine("Printers on this Network: ");
                    Console.WriteLine();
                    for (int i = 0; i < printQueueArray.Length; i++)
                    {
                        PrintQueue printer = printQueueArray[i];
                        //if printers are found on the network they appear here
                        //Console.WriteLine($"Printer Name: {printer.Name}");
                        //Console.WriteLine($"Printer Server: {printer.HostingPrintServer.Name}");
                        //System.Threading.Thread.Sleep(1000);
                        Console.WriteLine();
                    }
                    break;
                case 2:
                    //Console.WriteLine("Please Type out Printer Name you wish to Connect with: " + " (** Exact Spelling Needed **)\n");
                    //Console.Write("Enter it Here: ");
                    Mutex mutex = new Mutex();
                    if (mutex.WaitOne())
                    {
                        try
                        {
                            selectedPrinter = "Zan Image Printer (BW)";
                            bool printerExists = true;
                            bool openOnce = true;
                            /*
                                        for (int i = 0; i < printQueueArray.Length; i++)
                                        {
                                            PrintQueue printer = printQueueArray[i];
                                            //Console.WriteLine($"Printer Found: {printer.Name}\n");
                                            //System.Threading.Thread.Sleep(1000);
                                            for (int j = 0; j < printQueueArray.Length; j++)
                                            {
                                                if (string.Equals(printer.Name, selectedPrinter))
                                                {
                                                printerExists = true;
                                                break;
                                                }
                                            }
                            */
                            if (printerExists)
                            {
                                if (selectedPrinter == ("Zan Image Printer (BW)"))
                                {
                                    //Console.WriteLine($"Printer Found: {printer.Name}\n");
                                    Console.WriteLine($"You selected'{selectedPrinter}' for interaction\n");
                                    Console.WriteLine("Please wait...\n");
                                    //System.Threading.Thread.Sleep(2000);

                                    //try
                                    //{ 
                                    bool isControlPanelOpen = isProcessRunning("Devices and Printers"); 
                                    if (!isControlPanelOpen)
                                    {
                                        if (openOnce == true)
                                        {
                                            openOnce = false;
                                            CloseProcessByName("Devices and Printers");

                                            Console.WriteLine("Opening Control Panel..\n");
                                            Console.WriteLine("Stand By..\n");
                                            System.Threading.Thread.Sleep(2000);
                                            DirectionCommands();
                                            Console.WriteLine("Control Panel Opened\n");
                                            System.Threading.Thread.Sleep(1000);
                                            OpenPrinterProperties(selectedPrinter);

                                            System.Threading.Thread.Sleep(1000);
                                            AutomationElement printerProperitesDialog = FindPrinterPropertiesDialog();
                                            if (printerProperitesDialog != null)
                                            {
                                                string className = printerProperitesDialog.Current.ClassName;

                                                Console.WriteLine("Switching to Save Tab \n");

                                                //Console.WriteLine($"Class Name: {className}");

                                                AutomationElement saveTabItem = FindTabItem(printerProperitesDialog);

                                                saveTabItem.SetFocus();

                                                if (saveTabItem != null)
                                                {
                                                    string tabName = saveTabItem.Current.Name;
                                                    //Console.WriteLine("System found Matching Tab: " + tabName);

                                                    PropertyCondition controlTypeCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                                                    PropertyCondition nameCondition = new PropertyCondition(AutomationElement.NameProperty, "Folder:");

                                                    AndCondition conditions = new AndCondition(controlTypeCondition, nameCondition);

                                                    AutomationElement folderPathElement = saveTabItem.FindFirst(TreeScope.Descendants, conditions);

                                                    if (folderPathElement != null)
                                                    {
                                                        string folderPath = folderPathElement.GetCurrentPropertyValue(ValuePattern.ValueProperty) as string;
                                                        Console.WriteLine("Folder Path: " + folderPath);
                                                        Console.WriteLine();

                                                        string UserInputFolderPath;

                                                        do
                                                        {
                                                            //Console.WriteLine("Would you like to Change Folder Path? Y/N");
                                                            //Console.Write("Choice: ");
                                                            UserInputFolderPath = "y";

                                                            UserInputFolderPath.ToLower();
                                                            if (UserInputFolderPath == "y")
                                                            {
                                                                //Console.Write("Please Enter the new path for the folder" + "** Please Note: This path needs to be exact! **");
                                                                string newPath = "\\\\scclerk04.stark.local\\PrintQueue\\";
                                                                folderPath = newPath;

                                                                try
                                                                {
                                                                    AutomationElement okButton = FindOKButton(printerProperitesDialog);

                                                                    if (okButton != null)
                                                                    {
                                                                        if (okButton.TryGetCurrentPattern(InvokePattern.Pattern, out object patternobject))
                                                                        {
                                                                            ValuePattern valuePattern = (ValuePattern)folderPathElement.GetCurrentPattern(ValuePattern.Pattern);
                                                                            valuePattern.SetValue(folderPath);

                                                                            InvokePattern invokePattern = (InvokePattern)patternobject;
                                                                            invokePattern.Invoke();

                                                                            Console.WriteLine("Ok button clicked\n");
                                                                            openOnce = false;
                                                                        }
                                                                    }
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    Console.WriteLine("Error: Can't call newPathClass" + ex.Message);
                                                                }

                                                                break;
                                                            }

                                                        } while (UserInputFolderPath != "n" || UserInputFolderPath != "y");
                                                    }
                                                    break;
                                                }
                                            }
                                            else
                                            {
                                                Console.WriteLine("Couldn't connect... Please Try Again.\n");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Panel already running");
                                    }
                                    /*}
                                      catch (Exception AccessDenied)
                                      {
                                        Console.WriteLine("Access Denied");
                                        Console.WriteLine("Can't Access Printer Preferences");
                                        Console.WriteLine(AccessDenied.Message);
                                      }*/
                                }
                                //}
                            }
                            /*
                            else
                            {
                                Console.WriteLine($"Sorry.. Printer: '{selectedPrinter}' doesn't exist");
                                Console.ReadKey();
                                Environment.Exit(0);
                            }
                            */
                        }
                        finally
                        {
                            mutex.ReleaseMutex();
                        }
                    }
                    
                    break;
                case 3:
                    Environment.Exit(0);
                    break;
                default:
                    break;
            }
            //} while (optionToNumberConverter != 3);

            Console.WriteLine("System Update Complete!\n");
            System.Threading.Thread.Sleep(1000);
            Console.WriteLine("Auto Closing Windows.");

            int hWnd = FindWindow(null, "Zan Image Printer (BW) Properties");

            System.Threading.Thread.Sleep(1000);

            if (hWnd != 0) { PostMessage(hWnd, WM_CLOSE, 0, 0); } else { Console.WriteLine("Printer Properties couldnt be closed. Please Close Manually!"); }

            System.Threading.Thread.Sleep(2000);
        }

        private static AutomationElement FindOKButton(AutomationElement folderPathElement)
        {
            Condition okButtonCondition = new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                new PropertyCondition(AutomationElement.NameProperty, "OK"));
            AutomationElement okButton = folderPathElement.FindFirst(TreeScope.Descendants, okButtonCondition);
            return okButton;
        }

        private static void OpenPrinterProperties(string printerName)
        {
            try
            {
                if (printerName != null)
                {
                    Mutex mutex = new Mutex();
                    if (mutex.WaitOne())
                    {
                        try
                        {
                            Console.WriteLine("Opening Printer Properties\n");
                            System.Threading.Thread.Sleep(3000);

                            Process.Start("rundll32.exe", $"printui.dll,PrintUIEntry /p /n\"{printerName}\"");

                            System.Threading.Thread.Sleep(4000);

                            OpenPrinterPreferences();
                        }
                        finally
                        {
                            mutex.ReleaseMutex();
                        }
                    }
                }
            }
            catch (Exception CantOpenProperties)
            {
                Console.WriteLine("Error Opening Printer Properties: " + CantOpenProperties.Message);
            }
        }

        private static void OpenPrinterPreferences()
        {
            try
            {
                AutomationElement printerDialogFound = FindPrinterPreferenceDialog();
                if (printerDialogFound != null)
                {
                    AutomationElement preferencesButton = FindPrinterPreferenceButton(printerDialogFound);
                    if (preferencesButton != null)
                    {
                        InvokePattern invokePattern = preferencesButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                        invokePattern.Invoke();
                    }
                }
            }
            catch (Exception CantOpenPreferences)
            {
                Console.WriteLine("Can't Find Preferences: " + CantOpenPreferences.Message);
            }
        }
        private static void DirectionCommands()
        {
            string processName = "control";
            string arguments = "printers";
            bool isDeviceAndPrintersOpened = IsPcocessingRunningStill(processName,arguments);

            if (!isDeviceAndPrintersOpened)
            {
                Process.Start(processName, arguments);
            }
            else
            {
                Console.WriteLine("Device and Printers is already running");
            }
            
            System.Threading.Thread.Sleep(4000);

        }

        private static bool IsPcocessingRunningStill(string processName, string arguments)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            foreach (Process process in processes)
            {
                string processCommandLine = GetCommandLine(process);
                if (processCommandLine.Contains(arguments))
                {
                    return true;
                }
            }
            return false;
        }

        private static string GetCommandLine(Process process)
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT CommandLine From Win32_Process WHERE ProcessID = " + process.Id))
            {
                using (ManagementObjectCollection objects = searcher.Get())
                {
                    return objects.Cast<ManagementBaseObject>().SingleOrDefault()?["CommandLine"]?.ToString();
                }
            }
        }

        private static void CloseProcessByName(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            foreach (Process process in processes)
            {
                process.CloseMainWindow();
                process.WaitForExit();
            }
        }

        private static bool isProcessRunning(string processName)
        {
            Process[] processes = Process.GetProcessesByName(processName);
            return processes.Length > 0;
        }

        private static AutomationElement FindPrinterPreferenceButton(AutomationElement printerDialogFound)
        {
            if (printerDialogFound != null)
            {
                Condition condition = new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button), new PropertyCondition(AutomationElement.NameProperty, "Preferences..."));
                AutomationElement preferenceButton = printerDialogFound.FindFirst(TreeScope.Descendants, condition);
                Console.WriteLine("Opening Printer Preferences\n");
                return preferenceButton;
            }
            return null;
        }
        private static AutomationElement FindPrinterPreferenceDialog()
        {
            PropertyCondition windowConition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            PropertyCondition titleConition = new PropertyCondition(AutomationElement.NameProperty, "Zan Image Printer (BW) Properties");

            AndCondition condition = new AndCondition(windowConition, titleConition);

            return AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
        }
        private static AutomationElement FindPrinterPropertiesDialog()
        {
            PropertyCondition windowConition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
            PropertyCondition titleConition = new PropertyCondition(AutomationElement.NameProperty, "Zan Image Printer (BW) Printing Preferences");

            AndCondition condition = new AndCondition(windowConition, titleConition);

            return AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
        }
        //helper function to find the save tab
        private static AutomationElement FindTabItem(AutomationElement parent)
        {
            PropertyCondition tabControlCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Tab);

            AutomationElement tabControl = parent.FindFirst(TreeScope.Descendants, tabControlCondition);

            if (tabControl != null)
            {
                foreach (AutomationElement tabItem in tabControl.FindAll(TreeScope.Children, Condition.TrueCondition))
                {
                    if (tabItem.Current.Name == "Save")
                    {
                        return tabItem;
                    }
                }
            }
            return null;
        }
    }
}
