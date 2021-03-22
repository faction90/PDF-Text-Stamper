using System.Windows.Forms;
using System;
using System.IO;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Threading.Tasks;

//TODO : READ FROM EXCEL
namespace PDF_Text_Stamper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
        }

        //Stamp text to PDF
        //TODO : add parameter for font, size, position
        private void StampPDF(PdfDocument pdfDoc, string project, string mepBy, string qty, string color)
        {
            int nbOfPages = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= nbOfPages; i++)
            {
                //Add text
                PdfCanvas canvas = new PdfCanvas(pdfDoc.GetPage(i));
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(0, 0).ShowText(project).EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(50, 50).ShowText(mepBy).EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(100, 100).ShowText(qty.ToString()).EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(150, 150).ShowText(color).EndText();
            }
            pdfDoc.Close();
        }

        //Get all file in the folder and stamp them
        //TODO : validate path
        //       Add proper sufix to file and/or create new directory
        private async void buttonStamp_Click(object sender, System.EventArgs e)
        {

            if (Directory.Exists(folderBrowserDialog1.SelectedPath))
            {
                string[] fileEntries = Directory.GetFiles(folderBrowserDialog1.SelectedPath);

                //Load info from excel file DATA.xlsx
                var file = new FileInfo(folderBrowserDialog1.SelectedPath + "\\Data.xlsx");
                List<SpInfoBase> infoToStamp = await LoadExcelFile(file);

                foreach (var sp in infoToStamp)
                {
                    string currentPdf = folderBrowserDialog1.SelectedPath + "\\" + sp.SpFileName + ".pdf";
                    string newPdf = folderBrowserDialog1.SelectedPath + "\\" + sp.SpFileName + "_M.pdf";

                    PdfDocument pdfDoc = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(newPdf));
                    StampPDF(pdfDoc, "FT21-001", "Julien Tremblay", sp.SpQuantity, sp.SpColor);
                    sp.SpStatus = "Stamped";
                }

                ////Stamp all PDF in folder
                //foreach (string fileName in fileEntries)
                //{
                //    if (Path.GetExtension(fileName).ToUpper() == ".PDF")
                //    {
                //        PdfDocument pdfDoc = new PdfDocument(new PdfReader(fileName), new PdfWriter(fileName + "_M.pdf"));
                //        StampPDF(pdfDoc,"FT21-001","Julien Tremblay",4,"Ral 1023");
                //    }
                //}
                infoStampStatus.Text = "Fini!";
            }
        }

        private static async Task<List<SpInfoBase>> LoadExcelFile(FileInfo file)
        {
            List<SpInfoBase> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            int row = 5;
            int col = 1;

            while (string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
            {
                SpInfoBase sp = new();
                sp.SpFileName = ws.Cells[row, col].Value.ToString();
                sp.SpQuantity = ws.Cells[row, col + 1].Value.ToString();
                sp.SpColor = ws.Cells[row, col + 2].Value.ToString();
                output.Add(sp);
                row += 1;
            }

            return output;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            infoStampStatus.Text = "";
        }

        //Select a folder
        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            infoStampStatus.Text = "";
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox3.Text = folderBrowserDialog1.SelectedPath;
                buttonStamp.Enabled = true;
            }
        }
    }
}