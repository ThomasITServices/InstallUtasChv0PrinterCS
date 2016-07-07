using System;
using System.Diagnostics;
using System.Media;
using System.Threading;
using System.IO;
using Microsoft.Win32;
using IWshRuntimeLibrary;

namespace ThomasITServices
{
    class VarForInstaller
    {

        private string currentPath = Environment.CurrentDirectory;
        private string fileName;
        private string arguments;
        private string appName;

        public string CurrentPath
        {
            get
            {
                return currentPath;
            }
            set
            {
                currentPath = value;
            }
        }
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }
        public string Arguments
        {
            get
            {
                return arguments;
            }
            set
            {
                arguments = value;
            }
        }
        public string AppName
        {
            get
            {
                return appName;
            }
            set
            {
                appName = value;
            }
        }

    }

    class Installer : VarForInstaller
    {
        public Installer() { }
        public Installer(string AppName, string FileName, string Arguments)
        {
            this.AppName = AppName;
            this.FileName = FileName;
            this.Arguments = Arguments;
            this.Start();
        }
        public Installer(string AppName,string SourcePath, string FileName, string Arguments)
        {
            this.AppName = AppName;
            this.CurrentPath = SourcePath;
            this.FileName = FileName;
            this.Arguments = Arguments;
            this.Start();
        }

        public void Start()
        {
            try
            {
                Process p = new Process();
                p.StartInfo.FileName = this.CurrentPath + @"\" + this.FileName;
                p.StartInfo.Arguments = this.Arguments;
                Console.WriteLine("Installing {0}", this.AppName);
                p.Start();
                p.WaitForExit();
                Console.WriteLine("     {0} is now installed!!", this.AppName);
                SystemSounds.Exclamation.Play();
                Thread.Sleep(5000);
            }
            catch (Exception e)
            {
                SystemSounds.Hand.Play();
                Console.WriteLine("Exception caught: {0}!!", e.Message);
                Console.WriteLine("Current Path: {0}", this.CurrentPath);
                Console.WriteLine("File Name Needed: {0}", this.FileName);
                Console.WriteLine("Make sure {0} is in the same directory as this app.", this.FileName);
                Console.Read();
            }

        }
    }

    class VarFromTo
    {
        string sourcePath;
        string destinationPath;
        string fileName;
        public bool IsDir;
        public string SourcePath
        {
            get
            {
                return sourcePath;
            }
            set
            {
                sourcePath = value;
            }
                
        }
        public string DestinationPath
        {
            get
            {
                return destinationPath;
            }
            set
            {
                destinationPath = value;
            }

        }
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }

        }
    }

    class CopyTo : VarFromTo
    {
        public CopyTo() { }
        public CopyTo(string SourcePath, string DestinationPath)
        {
            this.IsDir = true;
            this.SourcePath = SourcePath;
            this.DestinationPath = DestinationPath;
            this.Start();
        }
        public CopyTo(string SourcePath, string DestinationPath, string FileName)
        {
            this.IsDir = false;
            this.SourcePath = SourcePath;
            this.DestinationPath = DestinationPath;
            this.FileName = FileName;
            this.Start();
        }

        public void Start()
        {
            try
            {
                if (!Directory.Exists(this.DestinationPath))
                {
                    Directory.CreateDirectory(this.DestinationPath);
                }

                    Console.WriteLine("Trying To Copy Files From: {0}", this.SourcePath);
                    if (IsDir)
                    {
                        if (Directory.Exists(this.SourcePath))
                        {
                            string[] fileNames = Directory.GetFiles(this.SourcePath);
                            foreach (string s in fileNames)
                            {
                                this.FileName = Path.GetFileName(s);
                                string destFile = Path.Combine(this.DestinationPath, this.FileName);
                            System.IO.File.Copy(s, destFile, true);
                                Console.WriteLine("Copied File To: {0}", destFile);
                            }
                        Console.WriteLine("Completed Copying");
                        }
                        else
                        {
                            Console.WriteLine("Source path does not exist!");
                        }
                    }
                    else if (!IsDir)
                    {
                        if (Directory.Exists(this.SourcePath))
                        {
                            string sourceFile = Path.Combine(this.SourcePath, this.FileName);
                            string destFile = Path.Combine(this.DestinationPath, this.FileName);
                        System.IO.File.Copy(sourceFile, destFile, true);
                            Console.WriteLine("Copied File To: {0}", destFile);
                            Console.WriteLine("Completed Copying");
                        }
                    }                
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception caught: {0}!!", e.Message);
                Console.Read();
            }
        }
    }

    class DeleteFrom : VarFromTo
    {
        public DeleteFrom() { }
        public DeleteFrom(string SourcePath)
        {
            this.IsDir = true;
            this.SourcePath = SourcePath;
            this.Start();
        }
        public DeleteFrom(string SourcePath,string FileName)
        {
            this.IsDir = false;
            this.SourcePath = SourcePath;
            this.FileName = FileName;
            this.Start();
        }
        public void Start()
        {
            try
            {
                
                if (Directory.Exists(this.SourcePath))
                {
                    if (IsDir)
                    {

                        try
                        {
                            Console.WriteLine("Try To Delete: {0}", this.SourcePath);
                            Directory.Delete(this.SourcePath, true);
                            Console.WriteLine("Deleted Directory: {0}", this.SourcePath);
                        }
                        catch (IOException e)
                        {
                            Console.WriteLine(e.Message);
                            Console.Read();
                        }

                    }

                    else if (!IsDir)
                    {
                        try
                        {
                            string f = String.Format(@"{0}\{1}", this.SourcePath, this.FileName);

                            Console.WriteLine("Try To Delete: {0}", f);
                            if (System.IO.File.Exists(f))
                            {
                                System.IO.File.Delete(f);
                                Console.WriteLine("Deleted File: {0}", f);
                            }
                            else
                            {
                                string m = String.Format("Please Check Path: {0}", f);
                                IOException e = new IOException(m);
                                throw e;
                            }
                            
                        }
                        catch (IOException e)
                        {
                            Console.WriteLine(e.Message);
                            Console.Read();
                        }

                    }
                }
                else
                {
                    string m = String.Format("Please Check Path: {0}", this.SourcePath);
                    IOException e = new IOException(m);
                    throw e;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception caught: {0}!!", e.Message);
                Console.Read();
            }
        }
    }

    class VarForRegistry
    {
        string subKeyPath;
        string keyName;
        string keyValue;
        public string SubKeyPath
        {
            set
            {
                subKeyPath = value;
            }
            get
            {
                return subKeyPath;
            }
        }
        public string KeyName
        {
            set
            {
                keyName = value;
            }
            get
            {
                return keyName;
            }
        }
        public string KeyValue
        {
            set
            {
                keyValue = value;
            }
            get
            {
                return keyValue;
            }
        }
    }

    class AddRKLM : VarForRegistry
    {
        AddRKLM(string LMSubKey, string KeyName, string KeyValue)
        {
            this.SubKeyPath = LMSubKey;
            this.KeyName = KeyName;
            this.KeyValue = KeyValue;
        }
        public void Start()
        {
            RegistryKey subKey = Registry.LocalMachine.OpenSubKey(this.SubKeyPath, true);
            subKey.SetValue(this.KeyName, this.KeyValue);

        }
        public void GetKeyValue()
        {
            RegistryKey subKey = Registry.LocalMachine.OpenSubKey(this.SubKeyPath, false);
            this.KeyValue = subKey.GetValue(this.KeyName).ToString();
            Console.WriteLine("Key: {0} Value: {1}", this.KeyName, this.KeyValue);

        }
    }

    class AddShortCut
    {
        string targetPath;
        string iconLocation = @"%SystemRoot%\system32\imageres.dll,3";
        string shortCutPath;
        public string TargetPath
        {
            get
            {
                return this.targetPath;
            }
            set
            {
                this.targetPath = value;
            }

        }
        public string IconLocation
        {
            get
            {
                return this.iconLocation;
            }
            set
            {
                this.iconLocation = value;
            }

        }
        public string ShortCutPath
        {
            get
            {
                return this.shortCutPath;
            }
            set
            {
                this.shortCutPath = value;
            }

        }

        public AddShortCut() { }
        public AddShortCut(string Path, string LinkTo)
        {
            this.ShortCutPath = Path;
            this.TargetPath = LinkTo;
            
            this.Start();
        }
        public AddShortCut(string Path, string LinkTo, string IconPath)
        {
            this.ShortCutPath = Path;
            this.TargetPath = LinkTo;
            this.IconLocation = IconPath;
            this.Start();
        }

        public void Start()
        {
            try
            {
                WshShell shell = new WshShell();
                IWshShortcut link = (IWshShortcut)shell.CreateShortcut(this.ShortCutPath);
                link.TargetPath = this.TargetPath;
                link.WorkingDirectory = this.TargetPath;
                link.IconLocation = this.iconLocation;
                link.Save();
                Console.WriteLine(this.shortCutPath);
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        
    }
}
