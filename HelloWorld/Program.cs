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

                
                 
                Node n=docManClient.CreateFolder(ref otAuth, parentID, "Property Number 51", "This is another dummy folder", null);
                long objectID = n.ID;
                docManClient.CreateFolder(ref otAuth, objectID, "Property Number 52", "This is another dummy folder", null);
                Console.WriteLine(n.Name);
                Console.WriteLine(n.ID);
                Console.WriteLine(n.Nickname);
                Console.WriteLine(n.Metadata);
                Console.ReadLine();
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
        }
    }
}
