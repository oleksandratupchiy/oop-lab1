using System;
using System.Windows.Forms;
using SpreadsheetApp11.UI;


namespace SpreadsheetApp11
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}