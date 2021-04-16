using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using SolidEdgeDraft;
using SolidEdgePart;
using SolidEdgeCommunity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;




namespace SolidEdgeMacro
{



    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string folder = @"C:\Users\j.batlle\Desktop\macro testing\save";
            SolidEdgeFramework.Application application = null;
            SolidEdgeFramework.SolidEdgeDocument activeDocument = null;
            SolidEdgeFramework.Documents documentList = null;
            SolidEdgeDraft.DraftDocument activeDraftDocument = null;
            SolidEdgePart.PartDocument activePartDocument = null;



            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();



                // Attempt to connect to a running instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                SolidEdgeFramework.DocumentTypeConstants documentType = GetDocumentType(application.ActiveDocument);
                documentList = application.Documents;
                activeDocument = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
                //string savePath = folder + @"\" + activeDocument.FullName + ".stp";
                //Console.WriteLine(savePath);
                //activeDocument.SaveAs(savePath);
                foreach (SolidEdgeFramework.SolidEdgeDocument document in documentList) 
                {
                    string savePath = folder + @"\" + document.FullName + ".stp";
                    Console.WriteLine(savePath);
                    document.SaveAs(savePath);
                    savePath = folder + @"\" + document.FullName + ".dxf";
                    Console.WriteLine(savePath);
                    document.SaveAs(savePath);
                    
                }
                Console.WriteLine("Todo ha salido a pedir de Milhouse");
                /*switch (documentType)
                {
                    case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument:
                        Console.WriteLine("Grabbed draft document");
                        activeDraftDocument = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;



                        break;
                    case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
                        Console.WriteLine("Grabbed part document");
                        activePartDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;
                        //generate the document route



                        //SaveAsPar(activePartDocument, documentRoute);
                        string documentName = folder + @"\" + GetFileName(activePartDocument.Name);
                        string extension = "stp";
                        SaveAsExtension(activePartDocument, documentName, extension);

                        Console.WriteLine("Todo ha salido a pedir de Milhouse");
                        break;



                    default:
                        break;



                }*/

                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            finally
            {
                OleMessageFilter.Unregister();
            }



        }



        private static void ReadJsonFile(string jsonFileIn)
        {
            dynamic jsonFile = JsonConvert.DeserializeObject(File.ReadAllText(jsonFileIn));
            Console.WriteLine($"Folder: { jsonFile["folder"]}");
        }



        private static bool IsExtensionValid(string strExtension)
        {
            string[] validFileTypes = { "par", "psm", "asm", "dft", "pwd", "stp",  };



            for (int i = 0; i < validFileTypes.Length; i++)
            {
                if (strExtension.ToLower() == validFileTypes[i].ToLower())
                {
                    return true;
                }
            }
            return false;
        }
        private static SolidEdgeFramework.DocumentTypeConstants GetDocumentType(object obj)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)obj;
            return document.Type;
        }




        private static void SaveAsExtension(SolidEdgePart.PartDocument oDoc, string route, string extension)
        {
            if (IsExtensionValid(extension))
            {
                string sExpFile = System.IO.Path.ChangeExtension(oDoc.FullName, "." + extension);
                string fullRoute = route + sExpFile;
                Console.WriteLine("Saved As: " + fullRoute);
                //oDoc.SaveCopyAs(fullRoute);
                oDoc.SaveAs(fullRoute);

            }
            else
            {
                Console.WriteLine("invalid extension: " + extension);
            }
        }
        private static string GetFileName(string fileName)
        {
            int fileExtPos = fileName.LastIndexOf(".");
            int filePathPos = fileName.LastIndexOf(@"\");
            if (filePathPos < 0)
            {
                filePathPos = 0;
            }
            if (fileExtPos < 0)
            {
                fileExtPos = fileName.Length;
            }
            fileName = fileName.Substring(filePathPos, fileExtPos);



            return fileName.ToLower();
        }
    }
}
