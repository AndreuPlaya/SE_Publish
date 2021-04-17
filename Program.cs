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

            try
            {
                // See "Handling 'Application is Busy' and 'Call was Rejected By Callee' errors" topic.
                OleMessageFilter.Register();

                // Attempt to connect to a running instance of Solid Edge.
                application = (Application)Marshal.GetActiveObject("SolidEdge.Application");
                //get active document
                activeDocument = (SolidEdgeDocument)application.ActiveDocument;
                
                //execute different behaviour for different documet type
                switch (GetDocumentType(application.ActiveDocument)) //grab document type form active document
                {
                    case DocumentTypeConstants.igDraftDocument:
                        Console.WriteLine("Grabbed draft document");
                        SaveAsExtension(activeDocument, folder, "dxf"); //save the active document on the specified folder as dxf
                        SaveAsExtension(activeDocument, folder, "pdf"); //save the active document on the specified folder as pdf
                        DraftDocument activeDraft = (DraftDocument)application.ActiveDocument; //cast the active document as a draftDocument to access the model link
                        
                        foreach(ModelLink modelLink in activeDraft.ModelLinks) //loop for all model links found in the model link
                        {
                            if (GetDocumentType((SolidEdgeDocument)modelLink.ModelDocument) == DocumentTypeConstants.igPartDocument)
                            { 
                                SaveAsExtension((SolidEdgeDocument)modelLink.ModelDocument, folder, "stp"); // cast the individual modelLlink.ModelDocument as a model document and save it
                                break;
                            }

                            if (GetDocumentType((SolidEdgeDocument)modelLink.ModelDocument) == DocumentTypeConstants.igAssemblyDocument)
                            {
                                SolidEdgeDocument asmDocument = (SolidEdgeDocument)modelLink.ModelDocument;// cast the individual modelLlink.ModelDocument as a model document

                                if (asmDocument.Name.Contains("MPF")) // quick string check if contains the letters MPF
                                {
                                    Console.WriteLine("Found MPF named document: " + asmDocument.Name);
                                    SaveAsExtension((SolidEdgeDocument)modelLink.ModelDocument, folder, "stp"); //save the model link document on the specified folder as stp
                                    break;
                                }
                                else
                                {
                                    Console.WriteLine("Found an non MPF asembly document: " + asmDocument.Name); // found nothing here
                                }
                            }

                        }
                        break;
                    case DocumentTypeConstants.igPartDocument:
                        Console.WriteLine("Grabbed part document");
                        SaveAsExtension(activeDocument, folder, "stp"); // save the part document in the specified folder as stp
                        break;
                    default:
                        Console.WriteLine("No valid document"); // found nothing here
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
                OleMessageFilter.Unregister(); // unlink application after completion
            }
        }

        private static DocumentTypeConstants GetDocumentType(object obj) //simple method to handle the conversion and get the document type
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