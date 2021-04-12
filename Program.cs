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
            Documents seDocuments = null;

            //document variables declaration
            DraftDocument seDraftDocument = null;
            Sheet sheet = null;



            ReadJsonFile("data.json");

            try
            {
                //register
                OleMessageFilter.Register();

                // Connect to a running instance of Solid Edge
                seApplication = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");

                //get the dobuments
                seDocuments = seApplication.Documents;

                //get active document
                seDraftDocument = (DraftDocument)seApplication.ActiveDocument;


                string tempPath = "C://";
                string fileName = "patata123";
                string extension = "pepe";

                seDraftDocument.SaveAs(tempPath + fileName + "." + extension);

                if (seDraftDocument != null)
                {
                    // Get a reference to the active sheet.
                    sheet = seDraftDocument.ActiveSheet;

                    SaveFileDialog dialog = new SaveFileDialog();

                    // Set a default file name
                    dialog.FileName = System.IO.Path.ChangeExtension(sheet.Name, ".emf");
                    dialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.DesktopDirectory);
                    dialog.Filter = "Enhanced Metafile (*.emf)|*.emf";

                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        // Save the sheet as an EMF file.
                        sheet.SaveAsEnhancedMetafile(dialog.FileName);
                        
                        Console.WriteLine("Created '{0}'", dialog.FileName);
                    }
                }
                else
                {
                    throw new System.Exception("No active document.");
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
                seDraftDocument = null;
                seDocuments = null;
                seApplication = null;
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
    }
    }
        
  


