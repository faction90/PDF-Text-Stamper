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

namespace PDF_Text_Stamper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();
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

        private void buttonStamp_Click(object sender, System.EventArgs e)
        {
            if (Directory.Exists(folderBrowserDialog1.SelectedPath))
            {
                //Load info from excel file DATA.xlsx
                var file = new FileInfo(folderBrowserDialog1.SelectedPath + "\\Data.xlsx");
                List<List<String>> matrixInfo = LoadExcelToMatrix(file);

                StampInfo(matrixInfo, folderBrowserDialog1.SelectedPath);

                infoStampStatus.Text = "Fini!"; //Let user know stamp is complete
            }
        }

        // Stamp info to each PDF
        private static void StampInfo(List<List<string>> matrixInfo, string selectedPath)
        {
            for (int i = 1; i < matrixInfo.Count; i++)
            {
                //Crée le path pour le fichier matrice[i][0]
                string currentPdf = selectedPath + "\\" + matrixInfo[i][0] + ".pdf";
                //Crée le path pour le nouveau fichier matrice[i][0] avec Stamp
                string newPdf = selectedPath + "\\" + matrixInfo[i][0] + "_M.pdf";
                //check si matrice[i][0] est un fichier qui existe
                if (File.Exists(currentPdf))
                {
                    PdfDocument pdfDoc = new PdfDocument(new PdfReader(currentPdf), new PdfWriter(newPdf));

                    //loop toutes les pages
                    int nbOfPages = pdfDoc.GetNumberOfPages();
                    for (int j = 1; j <= nbOfPages; j++)
                    {
                        //Add text
                        PdfCanvas canvas = new(pdfDoc.GetPage(j));

                        int posY = 0;
                        int posYHeight = 15;
                        //loop print header [0]
                        for (int k = 1; k < matrixInfo[0].Count; k++)
                        {
                            canvas
                                .BeginText()
                                .SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12)
                                .MoveText(0, posY)
                                .ShowText(matrixInfo[0][k])
                                .EndText();
                            posY += posYHeight;
                        }

                        posY = 0;
                        int posX = 150;
                        //loop print info [i]
                        for (int m = 1; m < matrixInfo[i].Count; m++)
                        {
                            canvas
                                .BeginText()
                                .SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12)
                                .MoveText(posX, posY)
                                .ShowText(matrixInfo[i][m])
                                .EndText();
                            posY += posYHeight;
                        }

                        //stamp lines and rectangle ?
                        //Color magentaColor = new DeviceCmyk(0f, 1f, 0f, 0f);
                        //canvas
                        //.SetStrokeColor(magentaColor)
                        //.MoveTo(0, 0)
                        //.LineTo(36, 806)
                        //.LineTo(559, 36)
                        //.LineTo(100, 500)
                        //.ClosePathStroke();
                        //canvas.SetStrokeColor(magentaColor).Rectangle(100, 100, 200, 100).SetLineWidth(1).Stroke();
                    }
                    pdfDoc.Close();
                }
                //Fichier introuvable
                else MessageBox.Show(String.Format("Le fichier {0} n'existe pas!", matrixInfo[i][0]));
            }
        }

        //Load info from Excel
        private static List<List<string>> LoadExcelToMatrix(FileInfo file)
        {
            //Crée une matrice qui contient [row][col]
            List<List<String>> matrix = new();

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];

                //Valeur du début des données dans Excel (équivaut a "A4")
                int row = 4;
                int col = 1;

                //Valeur pour index de la matrice, car row débute pas a 0
                int indexRow = 0;

                //Loop pour chaque lignes qui ne sont pas vide
                while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                {

                    matrix.Add(new List<String>()); // Ajoute un element a la matrice
                    //Loop pour chaque colones qui ne sont pas vide dans la meme ligne
                    while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                    {
                        matrix[indexRow].Add(ws.Cells[row, col].Value.ToString()); // Ajoute un sous elemet a la matrice
                        col++;
                    }
                    col = 1;
                    indexRow++;
                    row++;
                }
            }
            return matrix;
        }

        /* //SaveStatus A REVOIR POUR L'INSTANT, SI FICHIER EXISTE PAS MESSAGEBOX EST AFFICHER
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
        */
    }
}