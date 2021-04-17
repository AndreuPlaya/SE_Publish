using System;
using System.Configuration;
using SolidEdgeFramework;
using SolidEdgeDraft;
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
            Application application = null;
            SolidEdgeDocument activeDocument = null;
            //DraftDocument activeDraft = null;

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();

                // Attempt to connect to a running instance of Solid Edge.
                application = (Application)Marshal.GetActiveObject("SolidEdge.Application");
                
                activeDocument = (SolidEdgeDocument)application.ActiveDocument;
                
                //execute different behaviour for different documet type
                switch (GetDocumentType(application.ActiveDocument)) //grab document type form active document
                {
                    case DocumentTypeConstants.igDraftDocument:
                        Console.WriteLine("Grabbed draft document");
                        //SaveAsExtension(activeDocument, folder, "dxf");
                        //SaveAsExtension(activeDocument, folder, "pdf");
                        DraftDocument activeDraft = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;
                        
                        foreach(ModelLink modelLink in activeDraft.ModelLinks)
                        {
                            if (GetDocumentType((SolidEdgeDocument)modelLink.ModelDocument) == DocumentTypeConstants.igPartDocument)
                            { 
                                SaveAsExtension((SolidEdgeDocument)modelLink.ModelDocument, folder, "stp");
                            }

                            if (GetDocumentType((SolidEdgeDocument)modelLink.ModelDocument) == DocumentTypeConstants.igAssemblyDocument)
                            {
                                
                                SolidEdgeDocument asmDocument = (SolidEdgeDocument)modelLink.ModelDocument;
                                Console.WriteLine("found an assembly document: " + asmDocument.Name);
                                if (asmDocument.Name.Contains("MPF"))
                                {
                                    SaveAsExtension((SolidEdgeDocument)modelLink.ModelDocument, folder, "stp");
                                }
                            }

                        }
                        break;
                    case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
                        Console.WriteLine("Grabbed part document");
                        SaveAsExtension(activeDocument, folder, "stp");
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

        private static DocumentTypeConstants GetDocumentType(object obj)
        {
            SolidEdgeDocument document = (SolidEdgeDocument)obj;
            return document.Type;
        }

        private static void SaveAsExtension(SolidEdgeDocument oDoc, string route, string extension)
        {
            string savePath = route + @"\" + System.IO.Path.ChangeExtension(oDoc.Name, "." + extension);
            oDoc.SaveCopyAs(savePath);
            Console.WriteLine("Saved As: " + savePath);
        }
    }
}