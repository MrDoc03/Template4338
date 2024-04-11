using Microsoft.Office.Interop.Excel;
using System;
using System.ComponentModel;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
//using Aspose.Cells;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls.Primitives;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Template4338

{
    /// <summary>
    /// Логика взаимодействия для WindowInfo.xaml
    /// </summary>
    public partial class _4338_Zverev : System.Windows.Window
    {
        public _4338_Zverev()
        {
            InitializeComponent();
        }

        static void Import()
        {
            // Создание нового приложения Excel
            string path = AppDomain.CurrentDomain.BaseDirectory;
            path += "2.xlsx";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // Подключение к базе данных с использованием EF
            using (var db = new MyDbContext())
            {
                Database database = db.Database;


                int numberOfRowDeleted = db.Database.ExecuteSqlCommand("Truncate table Tables");

                for (int i = 2; i <= rowCount; i++)
                {
                    // Создание нового объекта для каждой строки в Excel
                    Table myTable = new Table();
                    var descr = TypeDescriptor.GetProperties(myTable);
                    for (int j = 1; j <= colCount; j++)
                    {
                        // Новые строки Excel начинаются с 1, а не с 0
                        if ((xlRange.Cells[i, j] != null) && (xlRange.Cells[i, j].Value != null))
                        {

                            PropertyDescriptor property = descr[j];
                            switch (property.Name)
                            {
                                case "Id":; continue;

                                default:
                                    property.SetValue(myTable, System.Convert.ToString(xlRange.Cells[i, j].Value.ToString()));
                                    break;
                            }

                        }
                        
                    }
                    // Добавление объекта в контекст EF
                    db.Tables.Add(myTable);
                }

                // Сохранение изменений в базе данных
                db.SaveChanges();
                int numberOfRowDeleted3 = db.Database.ExecuteSqlCommand("DELETE FROM Tables WHERE (CodeOrder IS NULL)");

            }

            // Очистка
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Закрытие и освобождение
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }

        static void Export()
        {

            var excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
            Excel._Worksheet worksheet1 = workbook.Sheets[1];
            worksheet1.Name = "ExportedFromDatatable";
            worksheet1.Cells[1, 1] = "Вывод данных по времени проката: Код/Код заказа/Дата Создания/Код Клиента/Услуги";

            using (var db = new MyDbContext())
            {
                var data = db.Database.SqlQuery<ProkatTimeTable>("select distinct ProkatTime from Tables;");
                worksheet1.Cells[2, 1] = data.Count();
                int k = 1;
                foreach (var row in data)
                {

                    string query = "SELECT * FROM Tables where ProkatTime like N'%" + Convert.ToString(row.ProkatTime) + "%'";
                    Console.WriteLine(query);
                    var data2 = db.Database.SqlQuery<Table>(query);
                    int count = db.Tables.Count(p => p.ProkatTime.Contains(row.ProkatTime));
                    Excel.Worksheet worksheet = (Worksheet)workbook.Sheets.Add();
                    

                    int currentRow = 2; // Начинаем с первой строки
                    
                    foreach (var row2 in data2)
                    {
                        
                        worksheet.Cells[1, 1] = row2.ProkatTime;
                        worksheet.Cells[currentRow, 1] = row2.Id.ToString();
                        worksheet.Cells[currentRow, 2] = row2.CodeOrder.ToString();
                        worksheet.Cells[currentRow, 3] = row2.CreateDate.ToString();
                        worksheet.Cells[currentRow, 4] = row2.CodeClient.ToString();
                        worksheet.Cells[currentRow, 5] = row2.Services.ToString();
                        currentRow++; // Переходим к следующей строке
                    }

                    k++;
                }
            }
            
            string path = AppDomain.CurrentDomain.BaseDirectory;
            path += "4.xlsx";
            workbook.SaveAs(@path);
            workbook.Close();
            excelApp.Quit();
        }
        private void ImportClick(object sender, RoutedEventArgs e)
        {
            Import();
        }

        private void ExportClick(object sender, RoutedEventArgs e)
        {
            Export();
        }

        private void ImportJSON(object sender, RoutedEventArgs e)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            path += "2.json";

            // Чтение содержимого файла JSON в строку
            using (var db = new MyDbContext())
            {
                int numberOfRowDeleted = db.Database.ExecuteSqlCommand("Truncate table TableJSONs");
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    TableJSON tableJSON = new TableJSON();
                    var jsonTable = new TableJSON();
                    foreach (TableJSON jSON in JsonSerializer.Deserialize<TableJSON[]>(fs))
                    {
                        db.TablesJSON.Add(jSON);
                    };

                }
                // Сохранение изменений в базе данных
                db.SaveChanges();
            }


        }
        private void ExportWord(object sender, RoutedEventArgs e)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            using (var db = new MyDbContext())
            {
                var data = db.Database.SqlQuery<ProkatTimeTable>("select distinct ProkatTime from TableJSONs;");
                foreach (var row in data)
                {
                    
                    string query = "SELECT * FROM TableJSONs where ProkatTime like N'%" + Convert.ToString(row.ProkatTime) + "%'";
                    Console.WriteLine(query);
                    var data2 = db.Database.SqlQuery<TableJSON>(query);
                    int count = db.TablesJSON.Count(p => p.ProkatTime.Contains(row.ProkatTime));
                    object what = Word.WdGoToItem.wdGoToPage;
                    object which = Word.WdGoToDirection.wdGoToFirst;
                    object countpage = 1;
                    Word.Range startOfPageRange = wordDoc.GoTo(ref what, ref which, ref countpage);
                    
                    Word.Table wordTable = wordDoc.Tables.Add(wordApp.Selection.Range, count, 6);

                    int currentRow = 1; // Начинаем с первой строки
                    foreach (var row2 in data2)
                    {
                        wordTable.Cell(currentRow, 1).Range.Text = Convert.ToString(row2.Id);
                        wordTable.Cell(currentRow, 2).Range.Text = row2.CodeOrder;
                        wordTable.Cell(currentRow, 3).Range.Text = row2.CreateDate;
                        wordTable.Cell(currentRow, 4).Range.Text = row2.CodeClient;
                        wordTable.Cell(currentRow, 5).Range.Text = row2.Services;
                        wordTable.Cell(currentRow, 6).Range.Text = row2.ProkatTime;
                        currentRow++; // Переходим к следующей строке
                    }

                   

                    // Добавление разрыва страницы после таблицы
                    startOfPageRange.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
            }
            string path = AppDomain.CurrentDomain.BaseDirectory;
            path += "4.docx";
            wordDoc.SaveAs(@path);
            wordDoc.Close();
            wordApp.Quit();
        }
    }
}
