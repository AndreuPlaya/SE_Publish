using System;
using System.IO;
using System.Configuration;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using SolidEdgeDraft;
using SolidEdgePart;
using SolidEdgeCommunity;
using System.Runtime.InteropServices;






namespace SolidEdgeMacro
{

    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string folder = ConfigurationManager.AppSettings.Get("saveFolder");
            SolidEdgeFramework.Application application = null;
            SolidEdgeFramework.SolidEdgeDocument activeDocument = null;
            SolidEdgeFramework.Documents documentList = null;
            SolidEdgeDraft.DraftDocument activeDraft = null;
            SolidEdgeDraft.Sheet activeSheet = null;





            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();

                // Attempt to connect to a running instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                documentList = application.Documents;
                activeDocument = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
                
                //execute different behaviour for different documet type
                switch (GetDocumentType(application.ActiveDocument)) //grab document type form active document
                {
                    case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument:
                        Console.WriteLine("Grabbed draft document");
                        //SaveAsExtension(activeDocument, folder, "dxf");
                        //SaveAsExtension(activeDocument, folder, "pdf");
                        activeDraft = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;
                        activeSheet = activeDraft.ActiveSheet;
                        //activeSheet.DrawingObjects.Count
                        Console.WriteLine(activeSheet.DrawingObjects.Count);

                        Console.WriteLine(":>");
                        break;
                    case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
                        Console.WriteLine("Grabbed part document");
                        SaveAsExtension(activeDocument, folder, "stp");
                        Console.WriteLine(":>");
                        break;
                    default:
                        Console.WriteLine("No valid document");
                        break;
                }
                
                Console.WriteLine("Todo ha salido a pedir de Milhouse");
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

        private static SolidEdgeFramework.DocumentTypeConstants GetDocumentType(object obj)
        {
            SolidEdgeFramework.SolidEdgeDocument document = (SolidEdgeFramework.SolidEdgeDocument)obj;
            return document.Type;
        }

        private static void SaveAsExtension(SolidEdgeFramework.SolidEdgeDocument oDoc, string route, string extension)
        {
            string savePath = route + @"\" + System.IO.Path.ChangeExtension(oDoc.Name, "."+ extension);
            Console.WriteLine("Saved As: " + savePath);
            oDoc.SaveAs(savePath);
        }
    }
}