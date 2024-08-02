using QRCoder;
using DRAW = System.Drawing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using System.CodeDom.Compiler;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Text;
using System.IO;
using MathNet.Numerics.Distributions;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System.Runtime.InteropServices;

namespace QRCodeGenerate
{
    internal class Program
    {
        static void Main(string[] args)
        {
       
            string filePath = "D:\\googleAppProject\\中山大同區少年服務中心-動產明細清冊1120511.xlsx";
            string outPath = "D:\\QRCode\\";
            if (!Directory.Exists(outPath))
            {
                Directory.CreateDirectory(outPath);
            }
            
            var getDataRows = ReadExcel(filePath);
            for (int i = 0; i < getDataRows.Count(); i++)
            {
                GenerateQRCode("D:\\QRCode\\", getDataRows[i].Code + ".png", getDataRows[i].ShowName);
            }

            var htmlCode = GetQRCodeHtml(getDataRows);
            SaveQRCodePage(htmlCode);

            Console.ReadKey();
        }

        public static List<QRCodeSetting> ReadExcel(string completeFileName)
        {
            //初始化要存放資料的地方(此為泛字串陣列)
            var arrMyData = new List<QRCodeSetting>();


            XSSFWorkbook wk;
            XSSFSheet hst;
            XSSFRow hr;


            try
            {
                using (FileStream fs = new FileStream(completeFileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    wk = new XSSFWorkbook(fs);
                }
                //預設為第一張表  如要讀取其他工作表 請更改下面的值 0,1,2,3,4‧‧‧
                hst = (XSSFSheet)wk.GetSheetAt(0);
                #region Get Sheet Name
                //strSheetname = hst.SheetName;
                #endregion Get Sheet Name
                //取得標題
                hr = (XSSFRow)hst.GetRow(0);
                //取得欄位數量
                int dLastNum = hr.LastCellNum;
                //有多少筆資料 j=1 有標題,j=0 無標題 (從哪開始讀取資料)
                for (int j = 1; j <= hst.LastRowNum; j++)
                {
                    hr = (XSSFRow)hst.GetRow(j);
                    //宣告陣列 並將資料同時置入
                    var myData = new QRCodeSetting();
                    myData.Code = hr.GetCell(0) == null ? "" : hr.GetCell(0).ToString();  //如果儲存格資料為空白null
                    myData.ShowName = hr.GetCell(1) == null ? "" : hr.GetCell(1).ToString();

                    //將每一行的資料存放至arrMyData儲存 ，之後要做什麼應用取得 arrMyData 即可~
                    arrMyData.Add(myData);
                }
            }
            catch (Exception ex)
            {

            }

            return arrMyData;
        }

        public static void GenerateQRCode(string filePath, string fileName, string strCode)
        {
            QRCodeGenerator qrGenerator = new();

            QRCodeData qrCodeData = qrGenerator.CreateQrCode(strCode, QRCodeGenerator.ECCLevel.L);
            // 設定二維碼下方的標題

            //generator.Parameters.Barcode.CodeTextParameters.TwoDDisplayText = "HELLO";
            //generator.Parameters.CaptionBelow.Text = "ASPOSE";
            //generator.Parameters.CaptionBelow.Visible = true;
            //generator.Parameters.CaptionBelow.Font.Style = FontStyle.Bold;
            //generator.Parameters.CaptionBelow.Font.Size.Pixels = 18;
            //generator.Parameters.CaptionBelow.Font.FamilyName = "Verdana";

            PngByteQRCode qrCode = new PngByteQRCode(qrCodeData);

            byte[] qrCodeImage = qrCode.GetGraphic(20, DRAW.Color.FromArgb(255, 111, 0), DRAW.Color.FromArgb(43, 43, 43), true);
            //string outputFileName = @"Images\Code.png";
            using (MemoryStream memory = new MemoryStream(qrCodeImage))
            {
                string completeFile = GetCompleteFile(filePath, fileName);

                File.WriteAllBytes(completeFile, memory.ToArray());
                Console.WriteLine($"產生{completeFile}");

            }
        }

        public static string GetQRCodeHtml(List<QRCodeSetting> qrCodeList)
        {
            StringBuilder sb = new StringBuilder();


            sb.AppendLine(@"<html>																");
            sb.AppendLine(@"  <head>                                                              ");
            sb.AppendLine(@"    <meta charset=""utf-8"" />                                        ");
            sb.AppendLine(@"    <title>My test page</title>                                       ");
            sb.AppendLine(@"  </head>                                                             ");
            sb.AppendLine(@"  <body>		                                                        ");
            sb.AppendLine(@"	<table border=""1"">                                                  ");

            //刻列
            for (int i = 0; i < qrCodeList.Count(); i++)
            {
                if (i % 9 == 0)
                {
                    sb.AppendLine(@"	  <tr>                                                              ");
                }

                sb.AppendLine(@"		<td>                                                            ");
                sb.AppendLine(@$"			<img src='{qrCodeList[i].Code}.png'  width='100' height='100' /> ");
                sb.AppendLine(@"			<center>                                                    ");
                sb.AppendLine(@$"			<p>{qrCodeList[i].ShowName}</p>                                                 ");
                sb.AppendLine(@"			</center>                                                   ");
                sb.AppendLine(@"		</td>                                                           ");



                if (i % 9 == 8)
                {
                    sb.AppendLine(@"	  </tr>                                                              ");
                }

            }

            sb.AppendLine(@"	</table>                                                            ");
            sb.AppendLine(@"  </body>                                                             ");
            sb.AppendLine(@"</html>                                                               ");
            return sb.ToString();
        }

        public static void SaveQRCodePage(string htmlCode)
        {
            string completeFile = "D:\\QRCode\\QRCode清單.html";
            File.Delete(completeFile);
            StreamWriter sr = new StreamWriter(completeFile, true, Encoding.Default);
            sr.Write(htmlCode);
            sr.Flush();
            sr.Close();


        }

        private static string GetCompleteFile(string filePath, string fileName)
        {
            string completeFile = string.Format("{0}{1}", filePath, fileName);
            return completeFile;
        }
    }
}
