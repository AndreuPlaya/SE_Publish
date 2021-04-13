﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeCommunity.Extensions;
using SolidEdgeFramework;
using SolidEdgeDraft;
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
            //Main variables declaration
            SolidEdgeFramework.Application seApplication = null;

            try
            {
                //register
                OleMessageFilter.Register();

                // Connect to a running instance of Solid Edge
                seApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

                //get the documents
                Documents seDocuments = seApplication.Documents;

                for (int i = 0; i < seDocuments.Count; i++)
                {
                    SolidEdgeDocument document = (SolidEdgeDocument)seDocuments.Item(i);
                    Console.WriteLine(GetDocumentType(document));
                }
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
     
        private static bool IsExtensionValid(string filename)
        {
            string[] validFileTypes = { "par", "psm", "asm", "dft", "pwd" };
            string strFileNameOnly = string.Empty;

            strFileNameOnly = System.IO.Path.GetFileName(filename);
            if (string.IsNullOrEmpty(strFileNameOnly))
                return false;
            if (System.IO.Path.HasExtension(strFileNameOnly))
            {
                //it has an extension now check to see if it is a valid SE one
                string strExtension = System.IO.Path.GetExtension(strFileNameOnly);
                for (int i = 0; i < validFileTypes.Length; i++)
                {
                    if (strExtension.ToLower() == "." + validFileTypes[i].ToLower())
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }
        private static string GetDocumentType(SolidEdgeFramework.SolidEdgeDocument document)
        {
            string type = null;
            switch (document.Type)
            {
                case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument:
                    type="Assembly Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument:
                    type = "Draft Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igPartDocument:
                    type = "Part Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument:
                    type = "SheetMetal Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igUnknownDocument:
                    type = "Unknown Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igWeldmentAssemblyDocument:
                    type = "Weldment Assembly Document";
                    break;
                case SolidEdgeFramework.DocumentTypeConstants.igWeldmentDocument:
                    type = "Weldment Document";
                    break;
            }
            return type;
        }
    }
}

        
  


