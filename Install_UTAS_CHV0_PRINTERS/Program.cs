using System;
using System.Threading;
using ThomasITServices;

namespace Install_UTAS_CHV0_PRINTERS
{
    class Program
    {
        static void Main(string[] args)
        {
            string linkName = "Add Printers UTAS Chula Vista Site";
            string startDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
            string printServerStartMenuLink = String.Format("{0}\\programs\\{1}.lnk", startDir, linkName);
            string printServerDesktopLink = String.Format("{0}\\{1}.lnk", deskDir, linkName);
            string printServerPath = @"\\chv0fp01";
            string iconPath = @"%SystemRoot%\system32\imageres.dll,46";

            AddShortCut makeShortCut = new AddShortCut();
            makeShortCut.TargetPath = printServerPath;
            makeShortCut.IconLocation = iconPath;

            makeShortCut.ShortCutPath = printServerStartMenuLink;
            makeShortCut.Start();

            makeShortCut.ShortCutPath = printServerDesktopLink;
            makeShortCut.Start();

            Thread.Sleep(2000);

        }
    }
}
