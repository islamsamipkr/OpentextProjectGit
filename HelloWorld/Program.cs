using HelloWorld.CWS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace HelloWorld
{
    class Program
    {

        readonly char newlineSeparator = '\n';
        readonly char atTheRateSeparator = '@';
        readonly char colonSeparator = ':';
        readonly char dotSeparator = '.';
        readonly char tabSeparator = '\t';
        static void Main(string[] args)
        {
            // The user's credentials
            string username = "admin";
            string password = "livelink";

            // Create the Authentication service client
            AuthenticationClient authClient = new AuthenticationClient();

            // Store the authentication token
            string authToken = null;

            // Call the AuthenticateUser() method to get an authentication token
            try
            {
                Console.Write("Authenticating User...");
                authToken = authClient.AuthenticateUser(username, password);
                Console.WriteLine("Success!\n");
            }
            catch (FaultException e)
            {
                Console.WriteLine("Failed!");
                Console.WriteLine("{0} : {1}\n", e.Code.Name, e.Message);
                return;
            }
            finally
            {
                // Always close the client
                authClient.Close();
            }
            // Create the DocumentManagement service client
            DocumentManagementClient docManClient = new DocumentManagementClient();

            // Create the OTAuthentication object and set the authentication token
            OTAuthentication otAuth = new OTAuthentication();
            otAuth.AuthenticationToken = authToken;

            // Store the favorites
            Node[] favorites = null;

            // Call the GetAllFavorites() method to get the user's favorites
            try
            {
                Console.Write("Getting the user's favorites...");
                favorites = docManClient.GetAllFavorites(ref otAuth);
                Console.WriteLine("Success!\n");
            }
            catch (FaultException e)
            {
                Console.WriteLine("Failed!");
                Console.WriteLine("{0} : {1}\n", e.Code.Name, e.Message);
                return;
            }
            

            // Output the user's favorites
            Console.WriteLine("User's Favorites:\n");
            if (favorites != null)
            {
                foreach (Node node in favorites)
                {
                    Console.WriteLine(node.Name);
                }
            }
            else
            {
                Console.WriteLine("No Favorites.");
            }
            Console.WriteLine();
            // The local file path of the file to upload
            string filePath = @"C:\Users\saislam\Downloads\opentext_logo.jpg";

            // The ID of the parent container to add the document to
            int parentID = 7058;

            // Store the information for the local file
            FileInfo fileInfo = null;

            try
            {
                fileInfo = new FileInfo(filePath);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0}\n", e.Message);
                return;
            }

            // Create the DocumentManagement service client
          

            // Create the OTAuthentication object and set the authentication token
            
            otAuth.AuthenticationToken = authToken;

            // Store the context ID for the upload
            string contextID = null;

            // Call the CreateDocumentContext() method to create the context ID
            try
            {
                Console.WriteLine("Generating context ID...");
                //  contextID = docManClient.CreateDocumentContext(ref otAuth, parentID, fileInfo.Name, null, false, null);
                Console.WriteLine("Success!\n");
                //Node N=docManClient.
                Console.WriteLine("Enter the file path");
                String path = Console.ReadLine();//"C:\Backup\Book2.xlsx". "ADLP IT Comm Information Architecture_8192019.xlsx"
                Console.WriteLine("Enter Sheet number");
                int sheetNum = Convert.ToInt32(Console.ReadLine());//1
                                                                   //Console.WriteLine("Enter column number : Last Level");
                                                                   //int endLColumn = Convert.ToInt32(Console.ReadLine());//8
                int endLColumn = 2;
                // Console.WriteLine("Enter Last row numebr");
                //int endLRow = Convert.ToInt32(Console.ReadLine());//55
                int endLRow = 50;

                Node n = docManClient.CreateFolder(ref otAuth, parentID, "Property Number 51", "This is another dummy folder", null);

                Console.WriteLine(n.Name);
                Console.WriteLine(n.ID);
                Console.WriteLine(n.Nickname);
                Console.WriteLine(n.Metadata);
                Console.ReadLine();
              
                Stack<string> ll = new Stack<string>();
                List<Dictionary<string, string>> folderDataList = new List<Dictionary<string, string>>();
                Program prg = new Program();
                string[,] arr = prg.readExcel(path, sheetNum, endLColumn, endLRow, ref folderDataList);
                
                String header = (String.Format("{0,-260}{1,-50}", "Folder Name", "Folder Path"));
                Console.WriteLine(header);
                String fullpath = "";
                prg.printPathRecursive(arr, 0, 0, ll, ref fullpath);
                
                long objectID = n.ID;
                n=n.Position(docManClient.CreateFolder(ref otAuth, objectID, "Property Number 52", "This is another dummy folder", null));

            }


            catch (FaultException e)
            {
                Console.WriteLine("{0} : {1}\n", e.Code.Name, e.Message);
                return;
            }
            finally
            {
                // Always close the client
                docManClient.Close();
            }
            // Create a file stream to upload the file with
            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0}\n", e.Message);
                return;
            }
            void printPathRecursive(string[,] arr, int row, int col, Stack<string> stk, ref String path)
            {
                //in case of last column
                if (row > arr.GetLength(0) - 1)
                {
                    //printPath(stk);
                    return;
                }


                if (!string.IsNullOrEmpty(arr[row, col]))
                {
                    //push into stack
                    stk.Push(arr[row, col]);
                    path = path + printPath(stk);
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
                            col = moveLeftSide(arr, row, col, ref stk, false, ref path);
                            if (col < arr.GetLength(1) - 1)
                            {
                                col = col + 1;

                            }
                            else
                            {
                                stk.Pop();
                            }
                            //move to digonal
                            printPathRecursive(arr, row + 1, col, stk, ref path);
                        }
                        else
                        {
                            return;
                        }

                    }

                }
                else
                {//in case of column is empty
                 //printPath(stk);
                    if (row < arr.GetLength(0))
                    {
                        col = moveLeftSide(arr, row, col, ref stk, true, ref path);
                        printPathRecursive(arr, row + 1, col + 1, stk, ref path);
                    }
                    else
                    {
                        return;
                    }
                }

            }
            string[,] readExcel(string path, int sheetNumber, int endLColumn, int endLRow, ref List<Dictionary<string, string>> permissionDataList)
            {

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
                    int startLColumn = 0;
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

                            if (!string.IsNullOrEmpty(cellValue) && cellValue.Equals("L1"))
                            {
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
                    for (int i = startLRow + 1; i <= endLRow; i++)
                    {
                        for (int j = startLColumn; j <= endLColumn; j++)
                        {
                            //  Console.WriteLine(i+" : "+ j);
                            string cellValue = (string)(excelWorksheet.Cells[i, j] as Excel.Range).Value;

                            arr[i - (startLRow + 1), j - startLColumn] = cellValue;

                        }
                    }
                    List<string> permissionList = new List<string>();
                    for (int k = endLColumn + 1; k < colCount; k++)
                    {
                        string cellValue = (string)(excelWorksheet.Cells[startLRow - 1, k] as Excel.Range).Value;
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
                            if (string.IsNullOrEmpty(permissionList.ElementAt(k)))
                            {
                                break;
                            }
                            hash.Add(permissionList.ElementAt(k), cellValue);
                            k++;

                        }
                        permissionDataList.Add(hash);
                    }

                    for (int i = 0; i < permissionDataList.Count; i++)
                    {

                        foreach (KeyValuePair<string, string> pair in permissionDataList.ElementAt(i))
                        {
                            Console.WriteLine("KEY: " + pair.Key + "VALUE: " + pair.Value);

                        }

                        Console.WriteLine("-----------------------------------------------------" + i);

                    }

                    excelWorkbook.Close();
                    excelApp.Quit();
                }
                return arr;
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

        private string[,] readExcel(string path, int sheetNum, int endLColumn, int endLRow, ref List<Dictionary<string, string>> permissionDataList)
        {
            throw new NotImplementedException();
        }

        private void printExcelFile(string[,] arr)
        {
            throw new NotImplementedException();
        }
    }
}
