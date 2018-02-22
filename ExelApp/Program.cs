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
        public class ExcelApp {
            public static Excel.Application excelApp = new Excel.Application();
            public static Excel.Workbook workBook =  excelApp.Workbooks.Add();
            public static Excel.Worksheet workSheet = (Excel.Worksheet) workBook.Worksheets.get_Item(1);
           
        }
        static void Info(int que,string info) {
            int row= 1;
            int column= que - 1 + que;
            int count = 0;
            string obr;
            while ((String.Compare(obr = ExcelApp.workSheet.Cells[row,column].Text, info, true) != 0) & (ExcelApp.workSheet.Cells[row, column].Text != "")) {
                // 1 2 3
                // 1 3 5 
                row++;
            }
            Console.WriteLine(ExcelApp.workSheet.Cells[row, column].Text);
            Console.WriteLine(row);
            if (String.Compare(obr = ExcelApp.workSheet.Cells[row, column].Text, info, true) == 0)
            {
                count = Convert.ToInt32(ExcelApp.workSheet.Cells[row, column+1].Text);
                count++;
                ExcelApp.workSheet.Cells[row, 2] = count;
            }
            else
            {
                //Console.WriteLine(String.Compare(ExcelApp.workSheet.Cells[row, 1].Text, info, false));
                ExcelApp.workSheet.Cells[row, column+1] = "1";
                ExcelApp.workSheet.Cells[row, column] = info;
                Console.WriteLine("Добавлен новый элемент");
                
            }

            Main(null);
        }
        static void Count(int row,string info) {
            //Строковые ф-ии изучить для сравнения
           
            Console.WriteLine("Сравниваем со словом яблоко");
            string obr = "яблоко";
            string slovo = Console.ReadLine();
            
           
            if (String.Compare(obr, slovo, true) == 0)
            {
                Console.WriteLine("Это яблоко");
            }
            else {
                Console.WriteLine("Это не яблоко");
            }

            Main(null);

        }
        public static void Main(string[] args)
        {
            try { 
            int count = 1;
            string bookS;
            Console.WriteLine("введите команду");
            string info = Console.ReadLine();
            switch (info) {

                case "опрос":
                    Console.WriteLine("введите номаер вопроса");
                    int que = Int32.Parse(Console.ReadLine());
                    Console.WriteLine("введите любимый фрукт");
                    info = Console.ReadLine();
                    Info(que, info);
                    break;
                case "начало":
                    Console.WriteLine("введите номаер вопроса");

                    try { que = Convert.ToInt32(Console.ReadLine()); }
                    catch {
                        Console.WriteLine("error");
                        Main(null);
                        break;

                    }
                    //que = Convert.ToInt32(Console.ReadLine());
                    Console.WriteLine("введите любимый фрукт");
                    info = Console.ReadLine();
                    Info(que, info);
                    break;
                case "проверка":
                    Count(1, "1");
                    break;
                case "выход":
                    count = 1;
                    bookS = "Книга" + count + ".xlsx";
                    ExcelApp.workBook.SaveAs(bookS);
                    ExcelApp.excelApp.Workbooks.Close();
                    Marshal.ReleaseComObject(ExcelApp.workBook);
                    ExcelApp.excelApp.Quit();
                    Process[] List;
                    List = Process.GetProcessesByName("EXCEL");
                    foreach (Process proc in List)
                    {
                        proc.Kill();
                    }
                    GC.Collect();
                    break;
                case "конец":

                    count = 1;
                    bookS = "Книга" + count + ".xlsx";
                    ExcelApp.workBook.Close(true);
                    ExcelApp.excelApp.Quit();
                    GC.Collect();
                    break;
                case "результат":
                    ExcelApp.excelApp.Visible = true;
                    ExcelApp.excelApp.UserControl = false;
                    Main(null);
                    break;
                default:
                    Console.WriteLine("неверная команда");
                    Console.WriteLine("начало или опрос- начало работы");
                    Console.WriteLine("результат - выводит текущую таблицу");
                    Console.WriteLine("выход или конец - сохраниет таблицу в папке с докупентами и выходит");
                    Console.WriteLine("в случае вылета почистить диспетчер задач от процессов экселя");
                    Main(null);
                    break;


            }
        }
        catch {
                Console.WriteLine("Что-то пошло не так, если при выходе вы не сохранили книгу, то сделайте это");
                Main(null);
            }
          
        }
    }
}
