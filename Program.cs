using System;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage.Streams;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Threading;

namespace Pdf2Png
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Pdf2Png.exe (nishidataku@yamaha-motor.co.jp)");
            Console.WriteLine("使い方 : Pdf2Png.exe <PDFファイルパス1> <PDFファイルパス2> ...");
            Console.WriteLine("TIPS : A0サイズなど大きいpdfはdpiに300程度を設定すると、拡大しても文字がつぶれません。");
            Console.WriteLine("TIPS : pdfファイルをexeファイルorショートカットにD&DでもOK。");
            //引数が0ならファイルダイアログでpdfファイルを出してあげる。
            if (args.Length == 0)
            {
                //ファイルダイアログを出すためにしょうがないからSTAにしとく。
                Thread.CurrentThread.SetApartmentState(ApartmentState.Unknown);
                Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                using (OpenFileDialog od = new OpenFileDialog())
                {
                    od.Title = "ファイルを開く";
                    od.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    od.Filter = "pdfファイル(*.pdf;*.PDF)|*.pdf;*.PDF";
                    od.FilterIndex = 1;
                    od.Multiselect = true;

                    // ダイアログを表示する
                    DialogResult result = od.ShowDialog();

                    // 選択後の判定
                    if (result == DialogResult.OK)
                    {
                        //「開く」ボタンクリック時の処理
                        args = od.FileNames;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        return;
                    }
                }
            }

            //dpi指定
            ushort dpi = 96;
            Console.WriteLine("dpiを指定してください。(エンターでデフォルト値96を使用します。A0だと300程度が最適です。) : ");
            string dpi_string=Console.ReadLine();
            try
            {
                dpi = ushort.Parse(dpi_string);
            }
            catch (FormatException)
            {
                if (dpi_string == "")
                {
                    //何もせずデフォルト値を使う。
                }
                else
                {
                    Console.WriteLine("dpiは正の整数である必要があります。デフォルト値を使用します");
                } 
            }
            catch(OverflowException){
                Console.WriteLine("dpiは0~32767で必要があります。デフォルト値を使用します。");
            }
            catch (ArgumentNullException)
            {
                //何もせずデフォルト値を使う。
            }

            foreach (string pdfPath in args)
            {
                if (!File.Exists(pdfPath) || Path.GetExtension(pdfPath).ToLower() != ".pdf")
                {
                    Console.WriteLine($"無効なPDFファイル: {pdfPath}");
                    continue;
                }
                await ConvertPdfToPngAsync(pdfPath, dpi);
            }
            Console.WriteLine("変換が完了しました。");
            Console.ReadKey();
        }

        static async Task ConvertPdfToPngAsync(string pdfPath, ushort dpi)
        {
            try
            {
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
                                PdfPageRenderOptions renderOptions = new PdfPageRenderOptions();
                                renderOptions.DestinationWidth = (uint)Math.Round(page.Dimensions.ArtBox.Width / 96.0 * dpi);
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
                Console.WriteLine($"{pdfPath}をスキップします。");
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
