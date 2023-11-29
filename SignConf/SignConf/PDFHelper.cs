using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Data;
using System.ComponentModel;
using Ghostscript.NET;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Advanced;

namespace SignConf
{
    public class PDFHelper
    {
        public static string GSPath
        {
            get
            {
                // 32-bit
                return System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath) + @"\GS\" + "gsdll32.dll";
            }
        }
        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height, PixelFormat.Format32bppArgb);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }
        public System.Drawing.Image TryAnomizar(byte[] binaryPdfData, int Page, string password, Image Sig, int SigPosX, int SigPosY, int PercentageRedux, string Folder, List<string> AditionalFiles)
        {

            Image newie = new Bitmap(16, 16);
            Ghostscript.NET.GhostscriptVersionInfo version = new Ghostscript.NET.GhostscriptVersionInfo(new Version(0, 0, 0), GSPath, string.Empty, Ghostscript.NET.GhostscriptLicense.GPL);
            iTextSharp.text.pdf.PdfReader.unethicalreading = true;
            if (PercentageRedux == 0)
            { PercentageRedux = 100; }
            int NewSigX = (Sig.Width * PercentageRedux /100);
            int NewSigY = (Sig.Height * PercentageRedux /100);
            Image NewSig = ResizeImage(Sig, NewSigX, NewSigY);
            iTextSharp.text.pdf.PdfReader pdfreader;
            using (var pdfDataStream = new MemoryStream(binaryPdfData))
            {
                if (string.IsNullOrWhiteSpace(password))
                {
                    pdfreader = new iTextSharp.text.pdf.PdfReader(pdfDataStream);
                }
                else
                {
                    byte[] newpassword = System.Text.ASCIIEncoding.ASCII.GetBytes(password);
                    //byte[] bytes = new byte[password.Length * sizeof(char)];
                    //System.Buffer.BlockCopy(password.ToCharArray(), 0, bytes, 0, bytes.Length);
                    //new System.Text.ASCIIEncoding().GetBytes(password));

                    pdfreader = new iTextSharp.text.pdf.PdfReader(pdfDataStream, newpassword);

                }

                using (MemoryStream ms = new MemoryStream())
                {
                    using (PdfStamper stamper = new PdfStamper(pdfreader, ms, '\0', false) { FormFlattening = true, FreeTextFlattening = true })
                    {
                        for (int page = 1; page <= pdfreader.NumberOfPages; page++)
                        {
                            if (Page == page)
                            {
                                int BitMapPositionX = NewSigX;
                                int BitMapPositionY = NewSigY;
                                int AbsolutePositionX = SigPosX;
                                int AbsolutePositionY = SigPosY;
                               
                                    

                                #region StamImage
                                if (BitMapPositionX == 0 && BitMapPositionY == 0 && AbsolutePositionX == 0 && AbsolutePositionY == 0)
                                { }
                                else
                                {
                                    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(NewSig,System.Drawing.Imaging.ImageFormat.Png);
                                    //Sets the position that the image needs to be placed (ie the location of the text to be removed)
                                    image.SetAbsolutePosition(AbsolutePositionX, AbsolutePositionY);
                                    //Adds the image to the output pdf
                                    var pdfContentByte = stamper.GetOverContent(page);
                                    
                                    pdfContentByte.AddImage(image);
                                    var gstate = new PdfGState { FillOpacity = 0.1f, StrokeOpacity = 0.3f };
                                    pdfContentByte.SaveState();
                                    pdfContentByte.SetGState(gstate);

 
                                    //stamper.GetUnderContent(page).AddImage(image, true);
                                }
                                
                                    #endregion
                            }
                        }
                        if (string.IsNullOrWhiteSpace(password) == false)
                        {
                            stamper.SetEncryption(null, null, iTextSharp.text.pdf.PdfWriter.ALLOW_PRINTING | iTextSharp.text.pdf.PdfWriter.ALLOW_MODIFY_CONTENTS, iTextSharp.text.pdf.PdfWriter.DO_NOT_ENCRYPT_METADATA);
                        }
                        
                        stamper.Writer.CloseStream = false;
                        stamper.Close();
                        if (Folder != "")
                        {
                            if (AditionalFiles.Count == 0)
                            {
                                System.IO.File.WriteAllBytes(Folder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf", ms.ToArray());
                            }
                            else
                            {
                                List<byte[]> MergeJob = new List<byte[]>();
                                MergeJob.Add(ms.ToArray());
                                foreach (string AditionalFile in AditionalFiles)
                                {
                                    if (System.IO.Path.GetExtension(AditionalFile).Contains("pdf"))
                                    {
                                        MergeJob.Add(System.IO.File.ReadAllBytes(AditionalFile));
                                    }
                                    else
                                    {
                                        Image NewImg = Bitmap.FromFile(AditionalFile);
                                        using (iTextSharp.text.Document doc = new iTextSharp.text.Document(new iTextSharp.text.Rectangle(NewImg.Width, NewImg.Height)))
                                        {
                                            using (MemoryStream ms2 = new MemoryStream())
                                            {
                                                using (PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, ms2))
                                                {
                                                    doc.Open();
                                                    try
                                                    {

                                                        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(AditionalFile);
                                                        jpg.Alignment = iTextSharp.text.Element.ALIGN_LEFT;
                                                        jpg.SetAbsolutePosition(0, 0);
                                                        jpg.ScaleAbsoluteHeight(doc.PageSize.Height);
                                                        jpg.ScaleAbsoluteWidth(doc.PageSize.Width);
                                                        doc.Add(jpg);
                                                    }
                                                    catch { }
                                                    finally { doc.Close(); MergeJob.Add(ms2.ToArray()); }
                                                }
                                            }
                                        }
                                    }
                                }
                                byte[] CompletePDF = MergePDF(MergeJob);
                                System.IO.File.WriteAllBytes(Folder + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf", CompletePDF);
                            }
                        }
                    }

                    using (var rasterizer = new Ghostscript.NET.Rasterizer.GhostscriptRasterizer())
                    {
                        rasterizer.Open(ms, version, true);

                        for (int i = 1; i <= rasterizer.PageCount; i++)
                        {
                            if (Page == i)
                            {
                                newie = rasterizer.GetPage(300, 300, i);
                            }
                        }
                    }
                }
            }
            return newie;
        }
        public byte[] MergePDF(List<byte[]> pdfFiles)
        {

            byte[] sucess = null;
            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    PdfSharp.Pdf.PdfDocument outputPDFDocument = new PdfSharp.Pdf.PdfDocument();
                    foreach (byte[] pdfFile in pdfFiles)
                    {

                        PdfSharp.Pdf.PdfDocument inputPDFDocument = CompatiblePdfReader.Open(pdfFile);
                        outputPDFDocument.Version = inputPDFDocument.Version;

                        foreach (PdfSharp.Pdf.PdfPage page in inputPDFDocument.Pages)
                        {

                            outputPDFDocument.AddPage(page);
                        }

                        inputPDFDocument.Dispose();
                    }

                    
                    outputPDFDocument.Save(ms, false);
                    sucess = ms.ToArray();

                    outputPDFDocument.Close();
                    outputPDFDocument.Dispose();
                }



            }
            catch
            {

            }
            return sucess;
        }
    }
    static public class CompatiblePdfReader
    {
        /// <summary>
        /// uses itextsharp 4.1.6 to convert any pdf to 1.4 compatible pdf, called instead of PdfReader.open
        /// </summary>
        static public PdfSharp.Pdf.PdfDocument Open(string pdfPath)
        {
            using (var fileStream = new FileStream(pdfPath, FileMode.Open, FileAccess.Read))
            {
                var len = (int)fileStream.Length;
                var fileArray = new Byte[len];
                fileStream.Read(fileArray, 0, len);
                fileStream.Close();

                return Open(fileArray);
            }
        }

        /// <summary>
        /// uses itextsharp 4.1.6 to convert any pdf to 1.4 compatible pdf, called instead of PdfReader.open
        /// </summary>
        static public PdfSharp.Pdf.PdfDocument Open(byte[] fileArray)
        {
            return Open(new MemoryStream(fileArray));
        }

        /// <summary>
        /// uses itextsharp 4.1.6 to convert any pdf to 1.4 compatible pdf, called instead of PdfReader.open
        /// </summary>
        static public PdfSharp.Pdf.PdfDocument Open(MemoryStream sourceStream)
        {
            PdfSharp.Pdf.PdfDocument outDoc;
            sourceStream.Position = 0;

            try
            {
                outDoc = PdfSharp.Pdf.IO.PdfReader.Open(sourceStream, PdfDocumentOpenMode.Import);
            }
            catch (PdfReaderException)
            {
                //workaround if pdfsharp doesn't support this pdf
                sourceStream.Position = 0;
                var outputStream = new MemoryStream();
                var reader = new iTextSharp.text.pdf.PdfReader(sourceStream);
                var pdfStamper = new iTextSharp.text.pdf.PdfStamper(reader, outputStream) { FormFlattening = true };
                pdfStamper.Writer.SetPdfVersion(iTextSharp.text.pdf.PdfWriter.PDF_VERSION_1_4);
                pdfStamper.Writer.CloseStream = false;
                pdfStamper.Close();

                outDoc = PdfSharp.Pdf.IO.PdfReader.Open(outputStream, PdfDocumentOpenMode.Import);
            }

            return outDoc;
        }
    }
}
