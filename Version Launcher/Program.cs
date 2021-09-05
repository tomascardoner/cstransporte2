using System;
using System.Windows.Forms;

namespace CSLauncher
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length == 0)
            {
                MessageBox.Show("Debe especificar el nombre del archivo INI como argumento de línea de comandos.", CardonerSistemas.My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (args.Length > 1)
            {
                MessageBox.Show("Sólo debe especificar 1 argumento de línea de comandos (nombre del archivo INI sin la extensión).", CardonerSistemas.My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (args[0].EndsWith(".ini", StringComparison.InvariantCultureIgnoreCase))
            {
                MessageBox.Show("Debe especificar el nombre del archivo INI sin la extensión.", CardonerSistemas.My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string iniFileName = args[0];
            string iniFileFullName = args[0] + ".ini";
            string iniFileFullPath = System.IO.Path.Combine(Application.StartupPath, iniFileFullName);

            CardonerSistemas.Configuration.IniFile.IniFile iniFile = new CardonerSistemas.Configuration.IniFile.IniFile();
            iniFile.Read(iniFileFullPath, false);
            CardonerSistemas.Configuration.IniFile.Key key = iniFile.GetKey("ExeFileName", "System");
            string exeFileName;
            if (key == null)
            {
                return;
            }
            exeFileName = key.Value;

            string exeFileFullPath = System.IO.Path.Combine(Application.StartupPath, exeFileName);
            if (System.IO.File.Exists(exeFileFullPath))
            {
                System.Diagnostics.Process.Start(exeFileFullPath, $"CONFIG={iniFileName}");
            }
        }
    }
}
