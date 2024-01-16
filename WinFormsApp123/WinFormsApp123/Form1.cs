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

                        var paragraph1 = document.InsertParagraph();
                        paragraph1.Append(label1.Text + " ");
                        paragraph1.Append(textBox2.Text + " ");
                        paragraph1.Alignment = Xceed.Document.NET.Alignment.center;

                        var txt = document.InsertParagraph();
                        txt.Append(label3.Text);
                        txt.Alignment = Xceed.Document.NET.Alignment.center;

                        var paragraph2 = document.InsertParagraph();
                        paragraph2.Append(label2.Text + " ");
                        paragraph2.Append(textBox3.Text + " ");
                        paragraph2.Append(label4.Text + " ");
                        paragraph2.Append(textBox4.Text + " ");
                        paragraph2.Append(label5.Text + " ");
                        paragraph2.Append(textBox5.Text + " ");
                        paragraph2.Append(label6.Text + " ");
                        paragraph2.Append(textBox6.Text + " ");
                        paragraph2.Append(label7.Text + " ");
                        paragraph2.Alignment = Xceed.Document.NET.Alignment.right;

                        var ttxtxt = document.InsertParagraph();
                        ttxtxt.Append(label8.Text);
                        ttxtxt.Alignment = Xceed.Document.NET.Alignment.center;

                        document.InsertParagraph(textBox7.Text);

                        document.InsertParagraph(label9.Text);


                        var paragraph3 = document.InsertParagraph();
                        paragraph3.Append(label10.Text + " ");
                        paragraph3.Append(textBox8.Text + " ");
                        paragraph3.Append(label11.Text + " ");
                        paragraph3.Append(textBox9.Text + " ");
                        paragraph3.Append(label12.Text + " ");

                        var paragraph4 = document.InsertParagraph();
                        paragraph4.Append(label13.Text + " ");
                        paragraph4.Append(textBox10.Text + " ");
                        paragraph4.Append(label14.Text + " ");
                        paragraph4.Append(textBox11.Text + " ");
                        paragraph4.Append(label15.Text + " ");
                        paragraph4.Append(textBox12.Text + " ");
                        paragraph4.Append(label16.Text + " ");

                        var paragraph5 = document.InsertParagraph();
                        paragraph5.Append(label17.Text + " ");
                        paragraph5.Append(textBox13.Text + " ");

                        var paragraph6 = document.InsertParagraph();
                        paragraph6.Append(label18.Text + " ");
                        paragraph6.Append(textBox14.Text + " ");

                        var paragraph7 = document.InsertParagraph();
                        paragraph7.Append(label19.Text + " ");
                        paragraph7.Alignment = Xceed.Document.NET.Alignment.left;


                        var paragraph7_1 = document.InsertParagraph();
                        paragraph7_1.Append(textBox15.Text + " ");
                        paragraph7_1.Append(textBox16.Text + " ");
                        paragraph7_1.Append(textBox17.Text + " ");
                        paragraph7_1.Alignment = Xceed.Document.NET.Alignment.right;


                        var paragraph8 = document.InsertParagraph();
                        paragraph8.Append(" " + label20.Text + " ");
                        paragraph8.Append(label21.Text + " ");
                        paragraph8.Append(label22.Text + " ");
                        paragraph8.Alignment = Xceed.Document.NET.Alignment.right;

                        document.InsertParagraph(label23.Text);

                        var paragraph9 = document.InsertParagraph();
                        paragraph9.Append(textBox18.Text + " ");
                        paragraph9.Append(textBox19.Text + " ");
                        paragraph9.Append(textBox20.Text + " ");
                        paragraph9.Alignment = Xceed.Document.NET.Alignment.right;

                        var paragraph10 = document.InsertParagraph();
                        paragraph10.Append(label24.Text + " ");
                        paragraph10.Append(label25.Text + " ");
                        paragraph10.Append(label26.Text + " ");
                        paragraph10.Alignment = Xceed.Document.NET.Alignment.right;

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