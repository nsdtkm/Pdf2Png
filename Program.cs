using System;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage.Streams;
using System.Drawing;
using System.Drawing.Imaging;

namespace Pdf2Png
{
    class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Pdf2Png.exe (nishidataku@yamaha-motor.co.jp)");
                Console.WriteLine("使い方 : Pdf2Png.exe <PDFファイルパス1> <PDFファイルパス2> ...");
                Console.WriteLine("TIPS :pdfファイルをexeファイルorショートカットにD&DでもOK。");
                Console.ReadKey();
                return;
            }

            foreach (string pdfPath in args)
            {
                if (!File.Exists(pdfPath) || Path.GetExtension(pdfPath).ToLower() != ".pdf")
                {
                    Console.WriteLine($"無効なPDFファイル: {pdfPath}");
                    continue;
                }

                //Console.Write($"変換開始: {pdfPath} \r");
                await ConvertPdfToPngAsync(pdfPath);
            }
            Console.WriteLine("変換完了");
            Console.ReadKey();
        }

        static async Task ConvertPdfToPngAsync(string pdfPath)
        {
            try
            {
                // StorageFile の代わりに System.IO.Stream を使用
                using (FileStream stream = new FileStream(pdfPath, FileMode.Open, FileAccess.Read))
                {
                    PdfDocument pdfDocument = await PdfDocument.LoadFromStreamAsync(stream.AsRandomAccessStream());

                    string outputDirectory = Path.GetDirectoryName(pdfPath);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfPath);

                    for (uint i = 0; i < pdfDocument.PageCount; i++)
                    {
                        using (PdfPage page = pdfDocument.GetPage(i))
                        {
                            string outputFilePath = Path.Combine(outputDirectory, $"{fileNameWithoutExt}_page{i + 1:D3}.png");

                            using (InMemoryRandomAccessStream memoryStream = new InMemoryRandomAccessStream())
                            {
                                await page.RenderToStreamAsync(memoryStream);
                                SaveStreamToPng(memoryStream, outputFilePath);
                                if (i== pdfDocument.PageCount-1){
                                    Console.WriteLine($"{pdfPath} : {i + 1}/{pdfDocument.PageCount}");
                                }
                                else
                                {
                                    Console.Write($"{pdfPath} : {i + 1}/{pdfDocument.PageCount}\r");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラー: {ex.Message}");
                Console.ReadKey();
            }
        }

        static void SaveStreamToPng(IRandomAccessStream stream, string outputFilePath)
        {
            using (Stream netStream = stream.AsStream())
            using (Bitmap bitmap = new Bitmap(netStream))
            {
                bitmap.Save(outputFilePath, ImageFormat.Png);
            }
        }
    }

}
