using System.Windows.Forms;
using System;
using System.IO;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;

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
        public void StampPDF(PdfDocument pdfDoc, int j)
        {
            //Add text
            PdfCanvas canvas = new PdfCanvas(pdfDoc.GetPage(j));
            canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(265, 597).ShowText(textBox1.Text).EndText();
            //Add text
            PdfCanvas canvas2 = new PdfCanvas(pdfDoc.GetPage(j));
            canvas2.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 12).MoveText(0, 0).ShowText(textBox2.Text).EndText();
        }

        //Get all file in the folder and stamp them
        //TODO : validate path
        //       Add proper sufix to file and/or create new directory
        private void buttonStamp_Click(object sender, System.EventArgs e)
        {
            string[] fileEntries = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
            foreach (string fileName in fileEntries)
            {
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(fileName), new PdfWriter(fileName+"_M.pdf"));

                int nbOfPages = pdfDoc.GetNumberOfPages();
                for (int i = 1; i <= nbOfPages; i++)
                {
                    StampPDF(pdfDoc, i);
                }
                pdfDoc.Close();
            }
            MessageBox.Show("Fini!");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //Select a folder
        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = textBox3.Text;
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox3.Text = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}