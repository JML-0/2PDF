using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Layout;
using IWshRuntimeLibrary;

namespace _2PDF
{
    internal static class Program
    {
        private static readonly string[] Exts = { ".JPG", ".PNG" };
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            //Application.Run(new Form1());

            string lnk = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Microsoft\\Windows\\SendTo\\2PDF.lnk";

            if (!System.IO.File.Exists(lnk))
            {
                Configure(lnk);
            }

            string[] args = Environment.GetCommandLineArgs();
            List<string> argsok = new();

            for (int i = 1; i < args.Length; i++)
            {
                if (Exts.Contains(Path.GetExtension(args[i].ToUpper())))
                {
                    argsok.Add(args[i]);
                }
            }

            if (argsok.Count == 0)
            {
                Environment.Exit(0);
            }
               

            Convert(argsok);

            Environment.Exit(0);
        }
        private static void Convert(List<string> args)
        {
            try
            {
                foreach (string s in args)
                {
                    ImageData imageData = ImageDataFactory.Create(s);
                    PdfDocument pdfDocument = new(new PdfWriter(Path.GetDirectoryName(s) + "\\" + Path.GetFileNameWithoutExtension(s) + ".pdf"));
                    Document document = new(pdfDocument);

                    iText.Layout.Element.Image image = new(imageData);
                    image.SetWidth(pdfDocument.GetDefaultPageSize().GetWidth() - 50);
                    image.SetAutoScaleHeight(true);

                    document.Add(image);
                    pdfDocument.Close();
                }

                MessageBox.Show("Convertion terminée !", "2PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch 
            {
                MessageBox.Show("Impossible de convertir en PDF, veuillez mettre à jour le logiciel.", "2PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void Configure(string lnk)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnk);

            shortcut.TargetPath = Application.StartupPath + @"\2PDF.exe";
            shortcut.Save();

            MessageBox.Show("Configuration terminée, logiciel prêt !", "2PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}