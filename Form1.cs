using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Word_to_PDF_Converter
{
    public partial class Form1 : Form
    {
        string docFileName = string.Empty, defaultPath;

        public Form1()
        {
            InitializeComponent();
            btnConvert.Enabled = false;
            defaultPath = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.Personal));
            defaultPath = Path.Combine(defaultPath, "Desktop");


        }

        private void Button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "PDF Files|*.pdf";
            saveDialog.Title = "PDF file to be saved";
            saveDialog.InitialDirectory = defaultPath;
            saveDialog.FileName = Path.GetFileNameWithoutExtension(docFileName) + ".pdf";
            if (saveDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(saveDialog.FileName))
            {
                string outPdfFile = saveDialog.FileName;
                try

                {
                    object readOnly = true;
                    _Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Visible = false;
                    Document oDoc = wordApp.Documents.Open(docFileName, System.Reflection.Missing.Value, readOnly);
                    oDoc.Activate();
                    object format = WdSaveFormat.wdFormatPDF;
                    oDoc.SaveAs2(outPdfFile, format);
                    wordApp.Quit();

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                btnConvert.Enabled = false;
                docFileName = string.Empty;
                MessageBox.Show("File successfully converted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.InitialDirectory = defaultPath;
            openDialog.Title = "Select file to be converted";
            openDialog.Filter = "Word 2007 Documents (*.docx)|*.docx|Word 97-2003 Documents (*.doc)|*.doc";
            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                docFileName = openDialog.FileName;
                btnConvert.Enabled = true;
            }
        }
    }
}
