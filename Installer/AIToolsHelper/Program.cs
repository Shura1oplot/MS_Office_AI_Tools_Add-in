using System;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

internal class Program
{
    private static void Main(string[] args)
    {
        bool is_installed = false;

        if (args.Length != 2)
            return;

        if (args[1] == "powerpoint")
        {
            var ppApp = new PowerPoint.Application();

            if (args[0] == "install")
            {
                Console.Out.WriteLine("Installing add-in for PowerPoint...");
                Console.Out.Flush();

                is_installed = false;

                foreach (PowerPoint.AddIn addIn in ppApp.AddIns)
                {
                    // Console.WriteLine(ad.Name);

                    if (addIn.Name == "AI Tools")
                    {
                        is_installed = true;
                        addIn.Registered = MsoTriState.msoTrue;
                        addIn.Loaded = MsoTriState.msoTrue;
                        break;
                    }
                }

                if (!is_installed)
                {
                    string fileName = Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.ApplicationData), "Microsoft", "AddIns", "AI Tools.ppam");
                    var addIn = ppApp.AddIns.Add(fileName);
                    addIn.Registered = MsoTriState.msoTrue;
                    addIn.Loaded = MsoTriState.msoTrue;
                }
            }
            else if (args[1] == "uninstall")
            {
                Console.Out.WriteLine("Uninstalling add-in for PowerPoint...");
                Console.Out.Flush();

                foreach (PowerPoint.AddIn addIn in ppApp.AddIns)
                {
                    if (addIn.Name == "AI Tools")
                    {
                        addIn.Registered = MsoTriState.msoFalse;
                        addIn.Loaded = MsoTriState.msoFalse;
                        ppApp.AddIns.Remove("AI Tools");
                        break;
                    }
                }
            }

            for (int i = 0; i < 5; i++)
            {
                try
                {
                    ppApp.Quit();
                    break;
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }
        }
        else if (args[1] == "excel")
        {
            var excelApp = new Excel.Application();

            if (args[0] == "install")
            {
                Console.Out.WriteLine("Installing add-in for Excel...");
                Console.Out.Flush();

                is_installed = false;

                foreach (Excel.AddIn addIn in excelApp.AddIns)
                {
                    // Console.WriteLine(ad.Name);

                    if (addIn.Name == "AI Tools.xlam")
                    {
                        is_installed = true;
                        addIn.Installed = true;
                        break;
                    }
                }

                if (!is_installed)
                {
                    string fileName = Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.ApplicationData), "Microsoft", "AddIns", "AI Tools.xlam");
                    var addIn = excelApp.AddIns.Add(fileName);
                    addIn.Installed = true;

                }
            }
            else if (args[0] == "uninstall")
            {
                Console.Out.WriteLine("Uninstalling add-in for Excel...");
                Console.Out.Flush();

                foreach (Excel.AddIn addIn in excelApp.AddIns)
                {
                    if (addIn.Name == "AI Tools.xlam")
                    {
                        addIn.Installed = false;
                        break;
                    }
                }
            }

            for (int i = 0; i < 5; i++)
            {
                try
                {
                    excelApp.Quit();
                    break;
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }
        }
        else if (args[1] == "word")
        {
            Word.Application wordApp = new Word.Application();

            if (args[0] == "install")
            {
                Console.Out.WriteLine("Installing add-in for Word...");
                Console.Out.Flush();

                is_installed = false;

                foreach (Word.AddIn addIn in wordApp.AddIns)
                {
                    // Console.WriteLine(ad.Name);

                    if (addIn.Name == "AI Tools.dotm")
                    {
                        is_installed = true;
                        addIn.Installed = true;
                        break;
                    }
                }

                if (!is_installed)
                {
                    string fileName = Path.Combine(Environment.GetFolderPath(
                        Environment.SpecialFolder.ApplicationData), "Microsoft", "Word", "STARTUP", "AI Tools.dotm");
                    var addIn = wordApp.AddIns.Add(fileName, true);
                    addIn.Installed = true;
                }
            }
            else if (args[0] == "uninstall")
            {
                Console.Out.WriteLine("Uninstalling add-in for Word...");
                Console.Out.Flush();

                foreach (Word.AddIn addIn in wordApp.AddIns)
                {
                    if (addIn.Name == "AI Tools.dotm")
                    {
                        addIn.Installed = false;
                        addIn.Delete();
                        break;
                    }
                }
            }

            for (int i = 0; i < 5; i++)
            {
                try
                {
                    wordApp.Quit();
                    break;
                }
                catch
                {
                    Thread.Sleep(1000);
                }
            }
        }
    }
}