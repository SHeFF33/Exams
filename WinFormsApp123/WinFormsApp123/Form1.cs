using System;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Reflection.Metadata;
using iTextSharp.text.pdf.parser;
using Xceed.Words.NET;

namespace WinFormsApp123
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {
               SaveFileDialog savefile = new SaveFileDialog();
            savefile.Filter = "Word Documents (*.docx)|*.docx";
            if (savefile.ShowDialog() == DialogResult.OK)
            {

                    string filePath = savefile.FileName;
                    using (DocX document = DocX.Create(filePath))
                    {
                        var zag = document.InsertParagraph();
                        zag.Append(textBox1.Text);
                        zag.Alignment = Xceed.Document.NET.Alignment.left;
                        document.Save();
                    }

                    textBox1.Clear();

            }

        }
        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}