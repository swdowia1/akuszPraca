using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArkuaszPraca
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("zaczynamy");
            string[] dni = { "po" ,"wto","śro","czw","pia","so","nd"};
            var date = DateTime.Now;
            var lastDayOfMonth = DateTime.DaysInMonth(date.Year, date.Month);
            int start = 5;
       
            using (ExcelPackage eksel=new ExcelPackage())
            {
                var ff = eksel.Workbook.Worksheets.Add("start");
                for (int i = 1; i <=lastDayOfMonth; i++)
                {
                    int pos = i % 7;
                    DateTime ds = new DateTime(date.Year, date.Month, i);
                    ff.Cells[start+i, 2].Value = i;
                    int day = ((int)ds.DayOfWeek == 0) ? 7 : (int)ds.DayOfWeek;
                    day--;
                    ff.Cells[start+i, 3].Value = dni[day];


                }
                ff.Cells["D"+(lastDayOfMonth+start+1).ToString()].Formula = "=SUM(D"+(start+1).ToString()+":D"+(lastDayOfMonth+start).ToString()+")";
                eksel.SaveAs(new System.IO.FileInfo("aa.xlsx"));
            }
        }
    }
}
