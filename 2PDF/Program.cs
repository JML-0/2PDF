using DinkToPdf;
using DocumentFormat.OpenXml.Packaging;
using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Layout;
using IWshRuntimeLibrary;
using OpenXmlPowerTools;
using System.Drawing.Imaging;
using System.Xml.Linq;

namespace _2PDF
{
    internal static class Program
    {
        private static readonly string[] Exts = { ".JPG", ".PNG", ".DOCX" };
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
                MessageBox.Show("Configuration terminée, logiciel prêt !", "2PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else // Actualise
            {
                System.IO.File.Delete(lnk);
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

        public static string ParseDOCX(FileInfo fileInfo)
        {
            try
            {
                byte[] byteArray = System.IO.File.ReadAllBytes(fileInfo.FullName);
                using (MemoryStream memoryStream = new())
                {
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    using (WordprocessingDocument wDoc =
                                                WordprocessingDocument.Open(memoryStream, true))
                    {
                        int imageCounter = 0;
                        var pageTitle = fileInfo.FullName;
                        var part = wDoc.CoreFilePropertiesPart;
                        if (part != null)
                            pageTitle = (string)part.GetXDocument()
                                                    .Descendants(DC.title)
                                                    .FirstOrDefault() ?? fileInfo.FullName;

                        WmlToHtmlConverterSettings settings = new()
                        {
                            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                            PageTitle = pageTitle,
                            FabricateCssClasses = true,
                            CssClassPrefix = "pt-",
                            RestrictToSupportedLanguages = false,
                            RestrictToSupportedNumberingFormats = false,
                            ImageHandler = imageInfo =>
                            {
                                ++imageCounter;
                                string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                                ImageFormat? imageFormat = null;
                                if (extension == "png") imageFormat = ImageFormat.Png;
                                else if (extension == "gif") imageFormat = ImageFormat.Gif;
                                else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                                else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                                else if (extension == "tiff")
                                {
                                    extension = "gif";
                                    imageFormat = ImageFormat.Gif;
                                }
                                else if (extension == "x-wmf")
                                {
                                    extension = "wmf";
                                    imageFormat = ImageFormat.Wmf;
                                }

                                if (imageFormat == null) return null;

                                string? base64 = null;
                                try
                                {
                                    using (MemoryStream ms = new())
                                    {
                                        imageInfo.Bitmap.Save(ms, imageFormat);
                                        var ba = ms.ToArray();
                                        base64 = System.Convert.ToBase64String(ba);
                                    }
                                }
                                catch (System.Runtime.InteropServices.ExternalException)
                                { return null; }

                                ImageFormat format = imageInfo.Bitmap.RawFormat;
                                ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders()
                                                            .First(c => c.FormatID == format.Guid);
                                string mimeType = codec.MimeType;

                                string imageSource =
                                        string.Format("data:{0};base64,{1}", mimeType, base64);

                                XElement img = new(Xhtml.img,
                                        new XAttribute(NoNamespace.src, imageSource),
                                        imageInfo.ImgStyleAttribute,
                                        imageInfo.AltText != null ?
                                            new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                                return img;
                            }
                        };

                        XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                        var html = new XDocument(new XDocumentType("html", null, null, null),
                                                                                    htmlElement);
                        var htmlString = html.ToString(SaveOptions.DisableFormatting);
                        return htmlString;
                    }
                }
            }
            catch
            {
                return "Le fichier est ouvert ou contient des données corrompues.";
            }
        }

        public static Uri FixUri(string brokenUri)
        {
            string newURI;
            if (brokenUri.Contains("mailto:"))
            {
                int mailToCount = "mailto:".Length;
                brokenUri = brokenUri.Remove(0, mailToCount);
                newURI = brokenUri;
            }
            else
            {
                newURI = " ";
            }
            return new Uri(newURI);
        }
        private static void Convert(List<string> args)
        {
            FileInfo? fileInfo = null;
            var converter = new BasicConverter(new PdfTools());
            string htmlText = string.Empty;

            try
            {
                foreach (string s in args)
                {
                    if (Path.GetExtension(s.ToUpper()) == ".DOCX")
                    {
                        fileInfo = new FileInfo(s);

                        try
                        {
                            htmlText = ParseDOCX(fileInfo);
                        }
                        catch (OpenXmlPackageException e)
                        {
                            if (e.ToString().Contains("Invalid Hyperlink"))
                            {
                                using (FileStream fs = new(fileInfo.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                                {
                                    UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                                }
                                htmlText = ParseDOCX(fileInfo);
                            }
                        }

                        var doc = new HtmlToPdfDocument()
                        {
                            GlobalSettings = {
                            ColorMode = DinkToPdf.ColorMode.Color,
                            Orientation = DinkToPdf.Orientation.Portrait,
                            PaperSize = PaperKind.A4,
                            Out = Path.GetDirectoryName(s) + "\\" + Path.GetFileNameWithoutExtension(s) + ".pdf",
                            },
                            Objects = {
                                new ObjectSettings() {
                                    HtmlContent = htmlText,
                                    WebSettings = { DefaultEncoding = "utf-8" }
                                }
                            }
                        };

                        converter.Convert(doc);
                    }
                    else
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
                    
                }

                MessageBox.Show("Convertion terminée !", "2PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show("Impossible de convertir en PDF, veuillez mettre à jour le logiciel en l'exécutant`.\n\n Erreur :" + e, "2PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void Configure(string lnk)
        {
            WshShell shell = new();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnk);

            shortcut.TargetPath = Application.StartupPath + @"\2PDF.exe";
            shortcut.Save();
        }
    }
}