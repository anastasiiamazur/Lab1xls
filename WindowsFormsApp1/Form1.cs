using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Path.GetExtension(openFileDialog1.FileName) == ".xlsx")
                    {
                        using (SpreadsheetDocument sourcePresentationDocument = SpreadsheetDocument.Open(openFileDialog1.FileName, true))
                        {
                            textBox1.Text += "Author:         " + sourcePresentationDocument.PackageProperties.Creator + Environment.NewLine;
                            textBox1.Text += "Title:          " + sourcePresentationDocument.PackageProperties.Title + Environment.NewLine;
                            textBox1.Text += "Subject:        " + sourcePresentationDocument.PackageProperties.Subject + Environment.NewLine;
                            textBox1.Text += "Category:       " + sourcePresentationDocument.PackageProperties.Category + Environment.NewLine;
                            textBox1.Text += "Keywords:       " + sourcePresentationDocument.PackageProperties.Keywords + Environment.NewLine;
                            textBox1.Text += "Description:    " + sourcePresentationDocument.PackageProperties.Description + Environment.NewLine;
                            textBox1.Text += "ContentStatus:  " + sourcePresentationDocument.PackageProperties.ContentStatus + Environment.NewLine;
                            textBox1.Text += "Revision:       " + sourcePresentationDocument.PackageProperties.Revision + Environment.NewLine;
                            textBox1.Text += "Created:        " + sourcePresentationDocument.PackageProperties.Created.ToString() + Environment.NewLine;
                            textBox1.Text += "Modified:       " + sourcePresentationDocument.PackageProperties.Modified.ToString() + Environment.NewLine;
                            textBox1.Text += "LastModifiedBy: " + sourcePresentationDocument.PackageProperties.LastModifiedBy + Environment.NewLine;

                        }
                    }
                    else textBox1.Text = "ERROR";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }
    }
}
