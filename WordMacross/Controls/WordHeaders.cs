
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMacross
{
    class WordHeaders : IHeader
    {
        Word.Application wordApp;
        public WordHeaders()
        {
            wordApp = new Word.Application(); // открытие Word
        }
        //Добавление верхнего колонтитула
        public void AddHeaderRange()
        {
            var files = new DirectoryInfo(ShowDialog()).GetFiles("*.docx"); //Получаем файлы с расширением *.docx
            Word.Document wordDoc = null;

            foreach (var item in files)
            {
                try
                {
                    wordDoc = Open(item.FullName); //получить путь
                    foreach (Word.Section section in wordDoc.Sections)
                    {
                        Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Font.Name = "Times New Roman";
                        headerRange.Font.Size = 12;
                        headerRange.Text = $"УП.03\t\tФилин Д.С.\n{DateTime.Now.ToShortDateString()}\t\tСмоленский М.С.";
                    }
                    wordDoc.SaveAs(item.FullName); //сохранить с заменой
                    Console.WriteLine(item.Name + " Файл изменён");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    wordApp.Quit();
                }
                finally 
                {
                    wordDoc.Close();
                }
            }
            wordApp.Quit(); // закрытие Word
        }
        //Диалоговое меню загрузки Word документов
        private string ShowDialog()
        {
            string path = string.Empty;

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                    path = fbd.SelectedPath;
            }
            return path;
        }
        //Открытие документа Word
        private Word.Document Open(object path)
        {
            return wordApp.Documents.Add(ref path);
        }
    }
}
