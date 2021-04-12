﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFramework;
using SolidEdgePart;
using SolidEdgeCommunity;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SolidEdgeMacro
{
    
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //Main variables declaration
            SolidEdgeFramework.Application seApplication = null;
            SolidEdgeFramework.Documents seDocuments = null;

            //document variables declaration
            SolidEdgePart.PartDocument sePartDocument = null;

            //file properties delaration
            SolidEdgeFramework.PropertySets propertySets = null; //Collection of all properties


            ReadJsonFile(@"C:\Users\DEPSOFTWARE02\source\repos\SolidEdgeMacro\bin\Debug\data.json");

            try
            {
                //register
                OleMessageFilter.Register();

                //connect to YOUR solidedge
                seApplication = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
                //get the dobuments
                seDocuments = seApplication.Documents;

                //get active document
                sePartDocument = (PartDocument)seApplication.ActiveDocument;

                //get collection og all properties
                propertySets = (PropertySets)sePartDocument.Properties;

                


            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw;
            }
            finally
            {
                sePartDocument = null;
                seDocuments = null;
                seApplication = null;
            }

        }
        private static void ReadJsonFile(string jsonFileIn)
        {
            dynamic jsonFile = JsonConvert.DeserializeObject(File.ReadAllText(jsonFileIn));
            Console.WriteLine($"Folder: { jsonFile["folder"]}");
        }
    }

    


}