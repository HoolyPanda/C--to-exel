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
            int column=1;
            int count = 0;

            //ExcelApp.workSheet.Cells[row, 3] =  "1";
            while ((ExcelApp.workSheet.Cells[row,1].Text != info) & (ExcelApp.workSheet.Cells[row, 1].Text != "")) {

                Console.WriteLine(row);
                row++;

            }
            if ((ExcelApp.workSheet.Cells[row, 1].Text == info))
            {
                
                count = Convert.ToInt32(ExcelApp.workSheet.Cells[row, 2].Text);
                count++;
                ExcelApp.workSheet.Cells[row, 2] = count;
            }
            else {
               // row++;
                ExcelApp.workSheet.Cells[row, 2] = "1";
                ExcelApp.workSheet.Cells[row, 1] = info;
                // ExcelApp.workSheet.Cells[row, 3] =  "11";
                //ExcelApp.workBook.SaveAs("Книга1.xlsx");
                // count = Convert.ToInt32 (ExcelApp.workSheet.Cells[row, 3].Text); 
                // row++;

                Console.WriteLine(ExcelApp.workSheet.Cells[row, 2].Text);
            }
           
            Main(null);



        }
        static void Count(int que,string answer) {
            if (Convert.ToString(ExcelApp.workSheet.Cells[2, 2]) == Convert.ToString(ExcelApp.workSheet.Cells[1, 1]) )
            {
                Console.WriteLine("OK");
                System.Threading.Thread.Sleep(2500);
            }
            else {
                Console.WriteLine("NO");
                System.Threading.Thread.Sleep(2500);
                ExcelApp.workSheet.Cells[3, 3] = ExcelApp.workSheet.Cells[2, 2];
            }
            ExcelApp.excelApp.Visible = true;
        }
        public static void Main(string[] args)
        {

            int count=1;
            string bookS;
            Console.WriteLine("введите команду");
            string info = Console.ReadLine();
            switch (info) {

                case "опрос" :
                    Console.WriteLine("введите номаер вопроса");
                    int que = 1;//Int32.Parse(Console.ReadLine());
                    Console.WriteLine("введите любимый фрукт");
                    info = Console.ReadLine();
                    Info(que, info);
                    break;
                case "начало" :
                    Console.WriteLine("введите номер вопроса");
                     que = 1;//Int32.Parse(Console.ReadLine());
                    Console.WriteLine("введите любимый фрукт");
                    info = Console.ReadLine();
                    Info(que, info);
                    break;
                case "выход":
                    count=1 ;
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
                    ExcelApp.excelApp.UserControl = true;
                    Main(null);
                    break;
                default:
                    Console.WriteLine("неверная команда");
                    Main(null);
                    break;


            }
        
          
        }
    }
}
