﻿using System;
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
                    case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument:
                        Console.WriteLine("Grabbed draft document");
                        SaveAsExtension(activeDocument, folder, "dxf");
                        SaveAsExtension(activeDocument, folder, "pdf");
                        DraftDocument activeDraft = (SolidEdgeDraft.DraftDocument)application.ActiveDocument;
                        
                        foreach(ModelLink modelLink in activeDraft.ModelLinks)
                        {
                            if (GetDocumentType(modelLink.ModelDocument) == DocumentTypeConstants.igPartDocument)
                            {
                                SaveAsExtension((SolidEdgeDocument)modelLink.ModelDocument, folder, "stp");
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
            string savePath = route + @"\" + System.IO.Path.ChangeExtension(oDoc.Name, "."+ extension);
            oDoc.SaveAs(savePath);
            Console.WriteLine("Saved As: " + savePath);
        }
    }
}