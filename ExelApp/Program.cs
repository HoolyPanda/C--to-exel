using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace ExelApp
{
    
    class Program
    {


       /* public interface getWorkbook
        {
            //void a();


            Excel.Workbook newWorkBook(Excel.Workbook value);

             Excel.Workbook oldWorkBook(Excel.Workbook value);
        }
        public class getNewWorkBook : getWorkbook
        {
         public Excel.Workbook newWorkBook(Excel.Workbook value)
         {
                 return ExcelApp.excelApp.Workbooks.Open("d:ПУть");
        }

            void getWorkbook.newWorkBook(Excel.Workbook value) {

            } 
            
           
        } */

        void IntTest() {


            
        }







        public class ExcelApp 
        {
       //     public static  Excel.Application excelApp = new Excel.Application(); 
         //   public static  Excel.Workbook workBook = excelApp.Workbooks.Open(@"C:\Users\Pavlo\Documents\КнигаМ.xlsx");
           // public static Excel.Worksheet workSheet = (Excel.Worksheet) workBook.Worksheets.get_Item(1);
           
        }

    
        static void Info(int que,string info, Excel.Application excelApp, Excel.Workbook workBook, Excel.Worksheet workSheet) {
            int row= 1;
            int column= que - 1 + que;
            int count = 0;
            string obr;
            while ((String.Compare(obr = workSheet.Cells[row,column].Text, info, true) != 0) & (workSheet.Cells[row, column].Text != "")) {
                // 1 2 3
                // 1 3 5 
                row++;
            }
            Console.WriteLine(workSheet.Cells[row, column].Text);
            Console.WriteLine(row);
            if (String.Compare(obr = workSheet.Cells[row, column].Text, info, true) == 0)
            {
                count = Convert.ToInt32(workSheet.Cells[row, column+1].Text);
                count++;
                workSheet.Cells[row, 2] = count;
            }
            else
            {
                //Console.WriteLine(String.Compare(ExcelApp.workSheet.Cells[row, 1].Text, info, false));
                workSheet.Cells[row, column+1] = "1";
                workSheet.Cells[row, column] = info;
                Console.WriteLine("Добавлен новый элемент");
                
            }

            MainWork(excelApp, workBook, workSheet, null);
        }

        static Excel.Workbook  OnStart(Excel.Application excelApp,Excel.Workbook workBook) {
            string bookS1;
            //int flag = 0;
            Console.WriteLine("Какой файл ищем?");
            bookS1 = Console.ReadLine();
            try
            {
                workBook = excelApp.Workbooks.Open(@"C:\Users\Pavlo\Documents\" + bookS1 + ".xlsx");
                Console.WriteLine("Продолжаем разговор");

            }
            catch
            {
                workBook = excelApp.Workbooks.Add();
                Console.WriteLine("Аааааать(");
            }
            return workBook;
            
        }
        static void Count(int row,string info, Excel.Worksheet workSheet) {
            //Строковые ф-ии изучить для сравнения
           // ExcelApp.workSheet.Cells[1, 1] = 1;
            int x=2;
            int y=2;
            var random = new Random();
            int val = random.Next(1, 11);
            for (y = 2; y != 14; y++) {
               // Console.WriteLine(y);
                //ExcelApp.workSheet.Cells[y, x-1] = "Пользователь"+(x-1);
                for (x = 2; x != 32; x++) {
                    workSheet.Cells[y,x] = val;
                    val = random.Next(1, 11);
                 //   Console.WriteLine(x);
                }

            }

            Main(null);

        }
        static void MainWork(Excel.Application  excelApp,Excel.Workbook  workBook, Excel.Worksheet workSheet, string info) {

            int count = 1;
            string bookS;
           
            //string info;
            if (info == null) {
                Console.WriteLine("введите команду");
                info = Console.ReadLine();
            }
            switch (info)
            {

                case "опрос":
                    Console.WriteLine("введите номаер вопроса");
                    int que = Int32.Parse(Console.ReadLine());
                    Console.WriteLine("введите любимый фрукт");
                    info = Console.ReadLine();
                    Info(que, info, excelApp, workBook, workSheet);
                    break;
                case "начало":
                    Console.WriteLine("введите номаер вопроса");
                    try { que = Convert.ToInt32(Console.ReadLine()); }

                    catch
                    {
                        Console.WriteLine("error");
                        MainWork(excelApp, workBook, workSheet, "начало");
                        break;
                    }
                    Console.WriteLine("введите любимый фрукт");
                    info = Console.ReadLine();
                    Info(que, info, excelApp, workBook, workSheet);
                    break;
                case "проверка":
                    Count(1, "1", workSheet);
                    break;
                /*   case "выход":
                       count = 1;
                       bookS = "КнигаM" + count + ".xlsx";
                       workBook.SaveAs(bookS);
                       excelApp.Workbooks.Close();
                       Marshal.ReleaseComObject(workBook);
                       excelApp.Quit();
                       Process[] List;
                       List = Process.GetProcessesByName("EXCEL");
                       foreach (Process proc in List)
                       {
                           proc.Kill();
                       }
                       GC.Collect();
                       break; */
                case "конец":
                    //   count = 1;

                    Console.WriteLine("Как назвать книгу?");
                    bookS = Console.ReadLine();
                 //   if (bookS == " ") {
                   //     bookS = "КнигаМ" ;
                    //}
                    //bookS = "КнигаМ"+ ".xlsx";
                    workBook.Close(true, bookS);
                    excelApp.Quit();
                    GC.Collect();
                    break;
                case "результат":
                    excelApp.Visible = true;
                    excelApp.UserControl = false;
                    MainWork(excelApp, workBook, workSheet, null);
                    break;
                default:
                    Console.WriteLine("неверная команда");
                    Console.WriteLine("начало или опрос- начало работы");
                    Console.WriteLine("результат - выводит текущую таблицу");
                    Console.WriteLine("выход или конец - сохраниет таблицу в папке с докупентами и выходит");
                    Console.WriteLine("в случае вылета почистить диспетчер задач от процессов экселя");
                   MainWork(excelApp ,workBook ,workSheet ,null);
                    break;


            }

            /*

            Console.WriteLine("Что-то пошло не так, если при выходе вы не сохранили книгу, то сделайте это");
            //int count = 1;
            //string bookS;
            count = 1;
            bookS = "Книга" + count + ".xlsx";
            workBook.SaveAs(bookS);
            excelApp.Workbooks.Close();
            Marshal.ReleaseComObject(workBook);
            excelApp.Quit();
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
            GC.Collect();
            */

        }
        public static void Main(string[] args)
        {
            //если флаг = 0, то вызываем три функции, которые  возвращают приложение, книгу и листок
            int flag = 0;
            //Convert.ChangeType(workbook , typeof(oldWorkBook));
            
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = null;

            if (flag == 0) {
                workBook = OnStart(excelApp,workBook);
                flag = 1;
            }

            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

            MainWork(excelApp ,workBook ,workSheet,null );
            

            
          
        }
    }
}
