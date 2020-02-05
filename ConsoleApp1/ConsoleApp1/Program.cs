using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;



namespace ConsoleApp1
{
    class Program
    {

        readonly char newlineSparator = '\n';
        readonly char atTheRateSparator = '@';
        readonly char colonSparator = ':';
        readonly char dotSparator = '.';
        readonly char tabSparator = '\t';
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the file path");
            String path=Console.ReadLine();//"C:\Backup\Book2.xlsx". "ADLP IT Comm Information Architecture_8192019.xlsx"
         

            Console.WriteLine("Enter Sheet number");
            int sheetNum = Convert.ToInt32(Console.ReadLine());//1



            //Console.WriteLine("Enter column number : Last Level");
            //int endLColumn = Convert.ToInt32(Console.ReadLine());//8
            int endLColumn = 12;
           // Console.WriteLine("Enter Last row numebr");
            //int endLRow = Convert.ToInt32(Console.ReadLine());//55
             int endLRow = 55;

            Console.WriteLine("Enter the file path");
            String out_path = Console.ReadLine();//"C:\Backup\Book_ouput.xlsx". "ADLP IT Comm Information Architecture_8192019.xlsx"Console.WriteLine("Enter the file path");
            

            Stack<string> ll = new Stack<string>();
            List<Dictionary<string, string>> permissionDataList = new List<Dictionary<string, string>>();
            Program prg = new Program();
            string[,] arr = prg.readExcel(path, sheetNum, endLColumn, endLRow,ref permissionDataList);
            prg.printExcelFile(arr);
            String header=(String.Format("{0,-260}{1,-50}", "Folder Name", "Folder Path"));
            Console.WriteLine(header);
            String fullpath = "";
            prg.printPathRecursive(arr,0,0,ll,ref fullpath);

            prg.writeToExcel(out_path, 1, fullpath, permissionDataList);//C:\Backup\Book_output.xls


            Console.ReadLine();
            Console.ReadLine();

       

        }

        string[,] readExcel(string path,int sheetNumber, int endLColumn,int endLRow, ref List<Dictionary<string, string>> permissionDataList) {

            Excel.Application excelApp = new Excel.Application();
            int rowCount = 0;
            int colCount = 0;
          
            string[,] arr = null;

            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[sheetNumber];
                Excel.Range excelRange = excelWorksheet.UsedRange;
                rowCount = excelRange.Rows.Count;
                colCount = excelRange.Columns.Count;

                //default values
                int startLColumn =0;
                int startLRow = 0;

                //user have to input these
                // endLColumn = 10;
                // endLRow = 55;
                Boolean isFound = false;


                //this for loop is used to get find the row and column from where level will start
                for (int i = 1; i <= rowCount && isFound == false; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //  Console.WriteLine(i+" : "+ j);
                        string cellValue = (string)(excelWorksheet.Cells[i, j] as Excel.Range).Value;
                        // Console.WriteLine(cellValue+";;;");

                        if (!string.IsNullOrEmpty(cellValue) && cellValue.Equals("L1")) {
                            startLColumn = j;
                            startLRow = i;
                            isFound = true;
                            break;
                        }

                    }
                }

                int rowIndex = endLRow - (startLRow);
                int colIndex = endLColumn - (startLColumn - 1);
                arr = new string[rowIndex, colIndex];
                for (int i = startLRow+1; i <= endLRow; i++)
                {
                    for (int j = startLColumn; j <= endLColumn; j++)
                    {
                        //  Console.WriteLine(i+" : "+ j);
                        string cellValue = (string)(excelWorksheet.Cells[i, j] as Excel.Range).Value;

                        arr[i- (startLRow + 1), j- startLColumn] = cellValue;

                    }
                }
                List<string> permissionList = new List<string>();
                for (int k = endLColumn + 1; k < colCount; k++) {
                    string cellValue = (string)(excelWorksheet.Cells[startLRow-1, k] as Excel.Range).Value;
                    permissionList.Add(cellValue);
                    Console.WriteLine(startLRow - 1 + " : " + k + cellValue);
                }

                
               
              
                for (int i = startLRow + 1; i <= endLRow; i++)
                {
                    int k = 0;
                    Dictionary<string, string> hash = new Dictionary<string, string>();
                    for (int j = endLColumn + 1; j < colCount; j++)
                    {
                        //  Console.WriteLine(i+" : "+ j);
                        string cellValue = (string)(excelWorksheet.Cells[i, j] as Excel.Range).Value;
                        if (string.IsNullOrEmpty(permissionList.ElementAt(k))) {
                            break;
                        }
                        hash.Add(permissionList.ElementAt(k), cellValue);
                        k++;

                    }
                    permissionDataList.Add(hash);
                }

                for (int i = 0; i < permissionDataList.Count; i++) {

                    foreach (KeyValuePair<string, string> pair in permissionDataList.ElementAt(i))
                    {
                        Console.WriteLine("KEY: " + pair.Key + "VALUE: " + pair.Value);
                     
                    }

                    Console.WriteLine("-----------------------------------------------------"+i);

                }

                    excelWorkbook.Close();
                excelApp.Quit();
            }
            return arr;
        }

        void printExcelFile(string [,]arr) {

            for (int i = 0; i < arr.GetLength(0); i++)
            {
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    if (string.IsNullOrEmpty(arr[i, j]))
                        Console.Write("." + tabSparator);
                    else
                        Console.Write(arr[i, j] + tabSparator);
                }
                Console.WriteLine();
            }
        }


        void printPathRecursive(string[,] arr, int row, int col, Stack<string> stk, ref String path)
        {
            //in case of last column
            if (row > arr.GetLength(0) - 1) {
                //printPath(stk);
                return;
            }


            if ( !string.IsNullOrEmpty(arr[row, col]))
            {
                //push into stack
                stk.Push(arr[row, col]);
                path=path+printPath(stk);
                // if cell having some data, move to digonal
                if (col < arr.GetLength(1) - 1 && !string.IsNullOrEmpty(arr[row + 1, col + 1]))
                {
                    //move to digonal
                    printPathRecursive(arr, row + 1, col + 1, stk, ref path);
                }
                else
                {
                    // printPath(stk);

                    
                   row = row + 1;
                   if (row < arr.GetLength(0))
                   {
                        col = moveLeftSide(arr, row, col, ref stk, false,ref path);
                        if (col < arr.GetLength(1) - 1)
                        {
                            col = col + 1;

                        }
                        else {
                            stk.Pop();
                        }
                        //move to digonal
                        printPathRecursive(arr, row + 1, col, stk,ref path);
                    }
                   else
                   {
                         return;
                   }
        
                }

            }
            else {//in case of column is empty
                  //printPath(stk);
                if (row < arr.GetLength(0))
                {
                    col = moveLeftSide(arr, row, col, ref stk, true,ref path);
                    printPathRecursive(arr, row + 1, col + 1, stk,ref path);
                }
                else
                {
                    return;
                }
            }

        }


        /**
         this method is to move  the cell to left side so that we can do backtracking
            and removed the older cell and update the new cell while traversing
             */
        private int moveLeftSide(string [,]arr,int row, int col, ref Stack<string> stk, bool isFlag, ref String path)
        {
            int index=col;
            if (isFlag == true)
            {
                index = col - 1;
            }
            for (int j = index; j >= 0; j--)
            {
                stk.Pop();
                if (!string.IsNullOrEmpty(arr[row, j]))
                {
                    stk.Push(arr[row, j]);
                    path = path + printPath(stk);
                    col = j;
                    break;
                }
            }
            return col;
        }

        /**
         This is method is used for print the stack in reverse order
             
             */
        private String printPath(Stack<string> stk)
        {
            StringBuilder sb = new StringBuilder();
            //sb.Append( String.Format("{0,-260}", stk.Peek()));
            sb.Append(stk.Peek()+ atTheRateSparator);
            //Console.Write(stk.Peek() + "\t\t");
            //from the top to bottom 
            for (int i = stk.Count - 1; i >= 0; i--)
            {
                //Console.Write(stk.ElementAt(i) + ":");
                if(!string.IsNullOrEmpty(stk.ElementAt(i)))
                    sb.Append(stk.ElementAt(i) + colonSparator);
            }
            sb.Remove(sb.Length - 1,1);
            sb.Append(newlineSparator);
            Console.Write(sb.ToString());
            return sb.ToString();
        }

        private void writeToExcel(String outputPath,int  sheetNum, string data_path, List<Dictionary<string, string>> permissionDataList)
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetNum);


            string[] lineArr = data_path.Split(newlineSparator);

            for (int i = 0; i < lineArr.Length-1; i++)
            {

                string[] row_col_data = lineArr[i].Split(atTheRateSparator);
                xlWorkSheet.Cells[i+1, 1] = row_col_data[0];
                xlWorkSheet.Cells[i+1, 2] = row_col_data[1];

                //StringBuilder sb = new StringBuilder();
                int k = 3;
                foreach (KeyValuePair<string, string> pair in permissionDataList.ElementAt(i))
                {

                    if (!string.IsNullOrEmpty(pair.Value)) {
                        xlWorkSheet.Cells[i + 1, k] = pair.Key + " -->" + pair.Value;
                        k++;
                        //sb.Append( pair.Key + " -->" + pair.Value + ",");
                    }
                    
                }
                //sb.Remove(sb.Length - 1, 1);

               // xlWorkSheet.Cells[i + 1, 3] = sb.ToString();
            }


            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(outputPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            Console.WriteLine("Excel file created , you can find the file");


        }
    }
}
