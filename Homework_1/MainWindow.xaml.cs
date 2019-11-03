using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Documents;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System.Reflection;


namespace Homework_1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_GetDiploma_Click(object sender, RoutedEventArgs e)
        {
            Word._Application application = null;
            Word._Document document = null;
            Object missingObj = Missing.Value;
            Object falseObj = false;

            try
            {
                string richText = new TextRange(StudentsTextBox.Document.ContentStart, StudentsTextBox.Document.ContentEnd).Text;
                List<Student> students = JsonConvert.DeserializeObject<List<Student>>(richText);

                //создаем обьект приложения word
                application = new Word.Application();
                // создаем путь к файлу
                Object templatePathObj = AppDomain.CurrentDomain.BaseDirectory + @"Diploma.dotx";

                int i = 1;
                foreach (var student in students)
                {
                    // если вылетим на этом этапе, приложение останется открытым
                    document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);

                    StringReplace(document, "%NAME%", student.Name);
                    StringReplace(document, "%STATUS%", student.Status);
                    StringReplace(document, "%GraduationDate%", student.GraduationDate.ToString("dd.MM.yyyy"));
                    StringReplace(document, "%GraduationPlace%", student.GraduationPlace);
                    Object pathToSaveObj = AppDomain.CurrentDomain.BaseDirectory + $@"student{i.ToString()}.docx";
                    document.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocumentDefault, ref missingObj,
                        ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                        ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj,
                        ref missingObj);
                    i++;


                    document.Close(ref falseObj, ref missingObj, ref missingObj);
                }

                application.Quit(ref missingObj, ref missingObj, ref missingObj);

                document = null;
                application = null;
                StudentsTextBox.Document.Blocks.Clear();
                StudentsTextBox.AppendText("Done");
            }
            catch (Exception exception)
            {
                document?.Close(ref falseObj, ref missingObj, ref missingObj);
                application?.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                StudentsTextBox.Document.Blocks.Clear();
                StudentsTextBox.AppendText(exception.Message);
            }
        }

        private void Button_GetStudents_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application objExcel = null;
            try
            {
                StudentsTextBox.Document.Blocks.Clear();
                objExcel = new Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Excel.Workbook objWorkBook = objExcel.Workbooks.Open(ExcelPathTextBox.Text, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                var objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];

                int i = 1;
                List<Student> students = new List<Student>();
                //Выбираем записи из столбца.
                while (!String.IsNullOrEmpty(objWorkSheet.Range["A"+i.ToString(), "A"+i.ToString()].Text.ToString()))
                {
                    students.Add(new Student
                    {
                        Name = objWorkSheet.Range["A" + i.ToString(), "A" + i.ToString()].Text.ToString(),
                        Status =objWorkSheet.Range["B" + i.ToString(), "B" + i.ToString()].Text.ToString(),
                        GraduationDate =
                            DateTime.Parse(objWorkSheet.Range["C" + i.ToString(), "C" + i.ToString()].Text.ToString()),
                        GraduationPlace = objWorkSheet.Range["D" + i.ToString(), "D" + i.ToString()].Text
                            .ToString()
                    });
                    i++;
                }
                objExcel.Workbooks.Close();
                objExcel.Quit();

                StudentsTextBox.AppendText(JsonConvert.SerializeObject(students));
                
            }
            catch (Exception exception)
            {
                objExcel?.Workbooks.Close();
                objExcel?.Quit();
                StudentsTextBox.Document.Blocks.Clear();
                StudentsTextBox.AppendText(exception.Message);
            }
        }

        private void StringReplace(Word._Document document, string strToFind, string replaceStr)
        {
            Object missingObj = Missing.Value;
            // обьектные строки для Word
            object strToFindObj = strToFind;
            object replaceStrObj = replaceStr;
            // диапазон документа Word
            Word.Range wordRange;
            //тип поиска и замены
            object replaceTypeObj = Word.WdReplace.wdReplaceAll;
            // обходим все разделы документа
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                // берем всю секцию диапазоном
                wordRange = document.Sections[i].Range;

                /*
                Обходим редкий глюк в Find, ПРИЗНАННЫЙ MICROSOFT, метод Execute на некоторых машинах вылетает с ошибкой "Заглушке переданы неправильные данные / Stub received bad data"  Подробности: http://support.microsoft.com/default.aspx?scid=kb;en-us;313104
                // выполняем метод поиска и  замены обьекта диапазона ворд
                wordRange.Find.Execute(ref strToFindObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref replaceStrObj, ref replaceTypeObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                */

                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };

                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }


        }
    }
}
