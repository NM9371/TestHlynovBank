using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;
using System.Linq;

namespace TestHlynovBank
{
    public partial class Form1 : Form
    {
        string[] samplePaths { get; set; }
        Word._Application oWord;
        Word._Document oDoc;
        BindingSource source = new BindingSource();

        private Dictionary<string, string> FillingData = new Dictionary<string, string>();
        public Form1()
        {
            InitializeComponent();
            samplePaths = Directory.GetFiles(Directory.GetCurrentDirectory() + @"\samples");
            string[] sampleNames = samplePaths.Select(path => Path.GetFileNameWithoutExtension(path)).ToArray();
            comboBox1.DataSource = sampleNames;

        }
        private void LoadSample(string i)
        {
            //CloseWord();
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            //создаем обьект приложения word
            oWord = new Word.Application();
            // создаем путь к файлу
            Object oTemplate = samplePaths[0];
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
            oWord.Visible = true;
            FillingData.Clear();
            foreach (dynamic bm in oDoc.Bookmarks)
            {
                dynamic bname = bm.Name;
                FillingData.Add(bname,"");
            }
            source.DataSource = FillingData;
            grid.DataSource = source;
            grid.Refresh();
        }
        private void FillingSample()
        {
            oWord.Visible = false;
            foreach (var fd in FillingData)
            {
                oDoc.Bookmarks[fd].Range.Text = FillingData.Values.ToString();
                oDoc.Bookmarks[fd].Delete();
            }
            oWord.Visible = true;
        }

        private void CloseWord()
        {
            oDoc.Close();
            oWord.Quit();
        }

        private void EnterButton_Click(object sender, EventArgs e)
        {
            FillingSample();
        }

        private void Form1_Closing(object sender, EventArgs e)
        {
            CloseWord();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSample(samplePaths[comboBox1.SelectedIndex]);
        }
    }
}
