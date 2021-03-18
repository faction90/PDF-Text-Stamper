using System.Windows.Forms;
using System;
using System.IO;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using System.Collections.Generic;

//TODO : READ FROM EXCEL
namespace PDF_Text_Stamper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        //Stamp text to PDF
        //TODO : add parameter for font, size, position
        private void StampPDF(PdfDocument pdfDoc)
        {
            int nbOfPages = pdfDoc.GetNumberOfPages();
            for (int i = 1; i <= nbOfPages; i++)
            {
                //Add text
                PdfCanvas canvas = new PdfCanvas(pdfDoc.GetPage(i));
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(265, 597).ShowText(textBox1.Text).EndText();
                //Add text
                PdfCanvas canvas2 = new PdfCanvas(pdfDoc.GetPage(i));
                canvas2.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(0, 0).ShowText(textBox2.Text).EndText();
            }
            pdfDoc.Close();
        }

        //Get all file in the folder and stamp them
        //TODO : validate path
        //       Add proper sufix to file and/or create new directory
        private void buttonStamp_Click(object sender, System.EventArgs e)
        {
            //---------------------------------------------------
            //List<SpInfoBase> items = new()
            //{
            //    new SpInfoBase{ SpFileName="TEST", SpQuantity=5, SpColor="Jaune"},
            //    new SpInfoBase{ SpFileName="TEST2", SpQuantity=1, SpColor="Bleu"}
            //};

            //MessageBox.Show(items[1].SpFileName);
            //---------------------------------------------------

            if (Directory.Exists(folderBrowserDialog1.SelectedPath))
            {
                string[] fileEntries = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
                foreach (string fileName in fileEntries)
                {
                    if (Path.GetExtension(fileName).ToUpper() == ".PDF")
                    {
                        PdfDocument pdfDoc = new PdfDocument(new PdfReader(fileName), new PdfWriter(fileName + "_M.pdf"));
                        StampPDF(pdfDoc);
                    }
                }
                MessageBox.Show("Fini!");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //Select a folder
        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox3.Text = folderBrowserDialog1.SelectedPath;
                buttonStamp.Enabled = true;
            }
        }
    }
}