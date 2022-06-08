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

namespace TestHlynovBank
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void FillingSample()
        {
            Word._Application oWord;
            Word._Document oDoc;
            Object oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            //создаем обьект приложения word
            oWord = new Word.Application();
            // создаем путь к файлу
            Object oTemplate = "D:\\Job\\C#\\TestHlynovBank\\samples\\sample2.dotx";
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);

            oDoc.Bookmarks["DataName"].Range.Text = NameTextBox.Text;
            oDoc.Bookmarks["DataName"].Delete();
            oDoc.Bookmarks["DataSurname"].Range.Text = SurnameTextBox.Text;
            oDoc.Bookmarks["DataSurname"].Delete();
            oWord.Visible = true;

        }

        private void EnterButton_Click(object sender, EventArgs e)
        {
            //File.Open("..\\samples\\a.txt", FileMode.Open);
            FillingSample();
        }
    }
}
