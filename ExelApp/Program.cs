using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;




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
             ExcelApp.workSheet.Cells[row, 3] =  "1";
            while ((ExcelApp.workSheet.Cells[row,1].Text != info)&& (ExcelApp.workSheet.Cells[row, 1].Text == info)) {
                Console.WriteLine(row);
                row++;

            }
            if ((ExcelApp.workSheet.Cells[row, 1].Text== info))
            {
                count = Convert.ToInt32(ExcelApp.workSheet.Cells[row, 3].Text);
                count++;
                ExcelApp.workSheet.Cells[row, 3] = count;
            }
            else {
                row++;
                ExcelApp.workSheet.Cells[row, 1] = info;
                // ExcelApp.workSheet.Cells[row, 3] =  "11";
                // count = Convert.ToInt32 (ExcelApp.workSheet.Cells[row, 3].Text); 
               // row++;

                Console.WriteLine(ExcelApp.workSheet.Cells[row, 3].Text);



            }
            ExcelApp.excelApp.Visible = true;
            ExcelApp.excelApp.UserControl = true;
            Main(null);
            ExcelApp.workBook.Close(true, "C:\\Price.xlsx");
          //  ExcelApp.excelApp.Quit();
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
           // ExcelApp.workSheet.Cells[2, 2] = "a" ;
            
            Console.WriteLine("Введите номер вопроса");
            int que = Int32.Parse(Console.ReadLine());
            Console.WriteLine("Какой ваш любимый Фрукт");

            string info = Console.ReadLine();
          //  ExcelApp.workSheet.Cells[1, 1] = info;
          
            Info(que,info);
        }
    }
}
