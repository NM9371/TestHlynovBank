using System;
using System.Data;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using Newtonsoft.Json;

namespace TestHlynovBank
{
    public partial class Form1 : Form
    {
        public static DataTable Bookmarks { get; set; } = new DataTable(); //Таблица с закладками и заполняемыми данными
        public static DataTable Samples { get; set; } = new DataTable(); //Таблица с названиями и путями шаблонов
        public static DataTable DataToSave { get; set; } = new DataTable(); //Таблица для записи логов
        public void SampleFolderCheck() //Метод для получения списка шаблонов 
        {
            foreach(string path in Directory.GetFiles(Directory.GetCurrentDirectory() + @"\samples"))
            {
                Samples.Rows.Add(Path.GetFileNameWithoutExtension(path), path);
            }

        }
        public void SaveData() //Метод для сохранения логов 
        {
            string DataPath = Directory.GetCurrentDirectory() + $"\\data\\{Samples.Rows[ComboBox1.SelectedIndex]["Name"].ToString()}.json";
            foreach (DataRow fd in Bookmarks.Rows)
            {
                DataToSave.Columns.Add(fd["Закладка"].ToString());
                DataToSave.Rows[0][fd["Закладка"].ToString()] = fd["Текст"].ToString();
            }

            if (File.Exists(DataPath)) //Если есть данные записаные ранее 
            {
                DataToSave.Merge(JsonConvert.DeserializeObject<DataTable>(File.ReadAllText(DataPath))); 
                //Объединить новые данные с хранящимися в файле
            }
            File.WriteAllText(DataPath,JsonConvert.SerializeObject(DataToSave)); //Сохраняем данные
        }
        public Form1()
        {
            InitializeComponent();
            DataColumn BookmarksColumn1 = new DataColumn();
            DataColumn BookmarksColumn2 = new DataColumn();
            BookmarksColumn1.DataType = System.Type.GetType("System.String");
            BookmarksColumn1.ColumnName = "Закладка";
            BookmarksColumn1.ReadOnly = true;
            BookmarksColumn1.Unique = true;
            Bookmarks.Columns.Add(BookmarksColumn1);
            BookmarksColumn2.DataType = System.Type.GetType("System.String");
            BookmarksColumn2.Unique = false;
            BookmarksColumn2.ColumnName = "Текст";
            BookmarksColumn2.ReadOnly = false;
            Bookmarks.Columns.Add(BookmarksColumn2);

            DataColumn SamplesColumn1 = new DataColumn();
            DataColumn SamplesColumn2 = new DataColumn();
            SamplesColumn1.DataType = System.Type.GetType("System.String");
            SamplesColumn1.ColumnName = "Name";
            SamplesColumn1.ReadOnly = true;
            SamplesColumn1.Unique = true;
            Samples.Columns.Add(SamplesColumn1);
            SamplesColumn2.DataType = System.Type.GetType("System.String");
            SamplesColumn2.Unique = true;
            SamplesColumn2.ColumnName = "Path";
            SamplesColumn2.ReadOnly = true;
            Samples.Columns.Add(SamplesColumn2);
            DataToSave.Rows.Add();

            SampleFolderCheck();
            //ComboBox1.DataSource = (from row in Samples.AsEnumerable()
            //                        select row.Field<string>("Name")).ToList<string>();
            ComboBox1.DataSource= Samples;
            ComboBox1.DisplayMember = "Samples";
            ComboBox1.ValueMember = "Name";

        }
        class WordSample //Класс работы с шаблоном
        {
            public static Word._Application oWord; //Процесс Word
            public static Word._Document oDoc; //Документ Word
            public static bool isWordOpen = false; //Переменная для закрытия процесса при перевыборе шаблона
            public static DataTable LoadSample(string SamplePath) //Метод для получения списка закладок выбранного шаблона 
            {
                CloseWord();
                isWordOpen = true;
                Object oMissing = System.Reflection.Missing.Value;
                Object oTrue = true;
                Object oFalse = false;
                oWord = new Word.Application();
                Object oTemplate = SamplePath;
                oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
                Bookmarks.Clear();
                foreach (dynamic bm in oDoc.Bookmarks)
                {
                    dynamic bname = bm.Name;
                    Bookmarks.Rows.Add(bname, "");
                }
                return Bookmarks;
            }
            public static void FillingSample() //Метод заполнения шаблона занесенными данными 
            {
                oWord.Visible = false;
                foreach (DataRow fd in Bookmarks.Rows)
                {
                    oDoc.Bookmarks[fd["Закладка"].ToString()].Range.Text = fd["Текст"].ToString();
                    oDoc.Bookmarks[fd["Закладка"].ToString()].Delete();
                }
                oWord.Visible = true;
            }
            public static void CloseWord() //Метод закрытия процесса Word 
            {
                if (isWordOpen)
                {
                    oDoc.Close(); //Закрытие документа
                    oWord.Quit(); //Закрытие процесса
                }
            }

        } 

        private void EnterButton_Click(object sender, EventArgs e)
        {
            WordSample.FillingSample();
            SaveData();
            WordSample.isWordOpen = false;
            Application.Exit();
        }

        private void Form1_Closing(object sender, EventArgs e)
        {
            WordSample.CloseWord();
        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            grid.DataSource = WordSample.LoadSample(Samples.Rows[ComboBox1.SelectedIndex]["Path"].ToString());
            grid.Columns[1].Width = 200;
            grid.Refresh();
            grid.Visible = true;
            EnterButton.Visible = true;
        }
    }
}
