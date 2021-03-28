using System.Windows.Forms;
using System;
using System.IO;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using System.Collections.Generic;
using OfficeOpenXml;
using iText.Kernel.Colors;

//-----------------------------------------------------------
//TODO add test for file exist (Excel and pdf)
//-----------------------------------------------------------

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
        private void StampPDF(PdfDocument pdfDoc, string project, string mepBy, string qty, string color)
        {
            int nbOfPages = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= nbOfPages; i++)
            {
                //-----------------------------------------------------------
                //TODO find the wright coordinate for the test to stamp
                //-----------------------------------------------------------
                //Add text
                PdfCanvas canvas = new PdfCanvas(pdfDoc.GetPage(i));
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(0, 0).ShowText(project).EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(50, 50).ShowText(mepBy).EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(100, 100).ShowText(qty.ToString()).EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(150, 150).ShowText(color).EndText();
                Color magentaColor = new DeviceCmyk(0f, 1f, 0f, 0f);
                //canvas
                //.SetStrokeColor(magentaColor)
                //.MoveTo(0, 0)
                ////.LineTo(36, 806)
                ////.LineTo(559, 36)
                //.LineTo(100, 500)
                //.ClosePathStroke();
                canvas.SetStrokeColor(magentaColor).Rectangle(100, 100, 200, 100).SetLineWidth(1).Stroke();
            }
            pdfDoc.Close();
        }

        private void buttonStamp_Click(object sender, System.EventArgs e)
        {
            if (Directory.Exists(folderBrowserDialog1.SelectedPath))
            {
                //-------------------------------------------------
                //TODO check if Data.xlsx exist in folder
                //string[] fileEntries = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                //-------------------------------------------------
                //Load info from excel file DATA.xlsx
                var file = new FileInfo(folderBrowserDialog1.SelectedPath + "\\Data.xlsx");
                //List<SpInfoBase> infoToStamp = LoadExcelFile(file);
                List<List<String>> matrixInfo = LoadExcelToMatrix(file);

                MessageBox.Show(matrixInfo[0][0]);
                /*
                //-------------------------------------------------
                //TODO check if pdf vs excel list exist, if not change sp.SpStatus = "File Not Found!!!";
                // make check not case sensitive?
                //-------------------------------------------------
                //Stamp each file found in excel
                foreach (var sp in infoToStamp)
                {
                    string currentPdf = folderBrowserDialog1.SelectedPath + "\\" + sp.SpFileName + ".pdf";
                    string newPdf = folderBrowserDialog1.SelectedPath + "\\" + sp.SpFileName + "_M.pdf";

                    PdfDocument pdfDoc = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(newPdf));
                    StampPDF(pdfDoc, sp.SpProject, sp.SpMepBy, sp.SpQuantity, sp.SpColor);
                    sp.SpStatus = "Stamped";
                }

                //Save Status To Data.xlsx
                SaveStatus(infoToStamp, file);

                //Let user know stamp is complete
                infoStampStatus.Text = "Fini!";
                */
            }
        }

        private static List<List<string>> LoadExcelToMatrix(FileInfo file)
        {
            List<List<String>> matrix = new(); //Creates new nested List

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];

                int row = 4;
                int col = 1;

                int indexRow = 0;

                while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                {
                    matrix.Add(new List<String>());

                    while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                    {
                        matrix[indexRow].Add(ws.Cells[row, col].Value.ToString());
                        col++;
                    }
                    col = 1;
                    indexRow++;
                    row++;
                }
            }
            return matrix;
        }

        private static List<SpInfoBase> LoadExcelFile(FileInfo file)
        {
            List<SpInfoBase> output = new();

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];

                int row = 5;
                int col = 1;

                while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                {
                    SpInfoBase sp = new();
                    sp.SpFileName = ws.Cells[row, col].Value.ToString();
                    sp.SpQuantity = ws.Cells[row, col + 1].Value.ToString();
                    sp.SpColor = ws.Cells[row, col + 2].Value.ToString();

                    sp.SpProject = ws.Cells[1, 2].Value.ToString();
                    sp.SpMepBy = ws.Cells[2, 2].Value.ToString();

                    output.Add(sp);
                    row += 1;
                }
            }
            return output;
        }

        private static void SaveStatus(List<SpInfoBase> infoToStamp, FileInfo file)
        {
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];

                int row = 5;
                int col = 4;

                foreach (var sp in infoToStamp)
                {
                    ws.Cells[row, col].Value = sp.SpStatus;
                    row++;
                }
                package.Save();
            }
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