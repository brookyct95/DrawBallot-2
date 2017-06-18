using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace DrawBallot
{
    static class Methods
    {
        private static Random rng = new Random();
        public static void Shuffle<T>(this IList<T> list)
        {
            int n = list.Count;
            while (n > 1)
            {
                n--;
                int k = rng.Next(n + 1);
                T value = list[k];
                list[k] = list[n];
                list[n] = value;
            }
        }
        public static string DataTableToCSV(this DataTable datatable, char seperator)
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataRow dr in datatable.Rows)
            {
                for (int i = 0; i < datatable.Columns.Count; i++)
                {
                    sb.Append(dr[i].ToString());

                    if (i < datatable.Columns.Count - 1)
                        sb.Append(seperator);
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        public static DataTable ConvertCSVtoDataTable1(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            DataTable dt = new DataTable();
            var tempvalue ="";

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("isDrawed", typeof(bool));
            int index = 0;
            while (!sr.EndOfStream)
            {
                tempvalue = sr.ReadLine();
                string[] rows = new string[tempvalue.Split(';').Length];
                index = 0;
                foreach (string value in tempvalue.Split(';'))
                {
                    rows[index] = value;
                    index++;
                }
                
                DataRow dr = dt.NewRow();
                string rowValue = null;
                for (int i = 0; i < rows.Length-1; i++)
                {       
                    rowValue += rows[i];
                    rowValue += '-';
                }
                rowValue = rowValue.Substring(0, rowValue.Length - 1);
                dr[0] = rowValue;
                dr[1] = ToBoolean(rows[rows.Length-1]);
                
                dt.Rows.Add(dr);
            }
            sr.Close();
            return dt;
        }

        public static DataTable ConvertCSVtoDataTable2(string strFilePath)
        {
            StreamReader sr = new StreamReader(strFilePath);
            
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("First Name");
            dt.Columns.Add("Last Name");
            dt.Columns.Add("Prize");
            dt.Columns.Add("Date");
            while (!sr.EndOfStream)
            {
                string[] rows = Regex.Split(sr.ReadLine(), ";");
                DataRow dr = dt.NewRow();
                for (int i = 0; i < 5; i++)
                {
                        dr[i] = rows[i];
                }
                dt.Rows.Add(dr);
            }
            sr.Close();
            return dt;
        }

        public static string ExcelToCSV(string FilePath, char Seperator)
        {
            StringBuilder sb = new StringBuilder();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(FilePath);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int RowCount = xlRange.Rows.Count;
            int ColCount = xlRange.Columns.Count;
            for (int i = 1; i <= RowCount; i++)
            {
                for (int j = 1; j <= ColCount; j++)
                {
                    sb.Append(xlRange.Cells[i,j].value2.ToString());
                    if (j < ColCount)
                        sb.Append(Seperator);
                }
                sb.Append("false");
                sb.AppendLine();
            }
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return sb.ToString();
        }
        public static bool ToBoolean(this string value)
        {
            switch (value.ToLower())
            {
                case "true":
                    return true;
                case "t":
                    return true;
                case "1":
                    return true;
                case "0":
                    return false;
                case "false":
                    return false;
                case "f":
                    return false;
                default:
                    throw new InvalidCastException("You can't cast a weird value to a bool!");
            }
        }
        public static ListViewItem CreateItem(string[] arr)
        {
            ListViewItem item = new ListViewItem(arr);
            return item;
        }
    }
    
}
