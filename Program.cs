﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFramework;
using SolidEdgePart;
using SolidEdgeCommunity;

namespace SolidEdgeMacro
{
    class Program
    {
        static void Main(string[] args)
        {
            //Main variables declaration
            SolidEdgeFramework.Application seApplication = null;
            SolidEdgeFramework.Documents seDocuments = null;

            //document variables declaration
            SolidEdgePart.PartDocument sePartDocument = null;

            //file properties delaration
            SolidEdgeFramework.PropertySets propertySets = null; //Collection of all properties

            //some properties summary delaration
            SolidEdgeFramework.Properties propertiesSummary = null; // collection of all summary properties
            SolidEdgeFramework.Property title = null;
            SolidEdgeFramework.Property subject = null;
            SolidEdgeFramework.Property author = null;
            SolidEdgeFramework.Property comments = null;

            //project properties declaration
            SolidEdgeFramework.Properties propertiesProject = null;
            SolidEdgeFramework.Property documentNumber = null;
            SolidEdgeFramework.Property revision = null;
            SolidEdgeFramework.Property projectName = null;

            //write variables
            string newTitle;
            string newSubject;
            string newAuthor = "patata123";
            string newcomments;

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

                //get collection of summary properties
                propertiesSummary = propertySets.Item("SummaryInformation");

                //get properties
                title = propertiesSummary.Item("Title");
                subject = propertiesSummary.Item("Subject");
                author = propertiesSummary.Item("Author");
                comments = propertiesSummary.Item("Comments");

                //Set properties
                title.set_Value(newAuthor);
                //
                //
                //


                //get project properties
                propertiesProject = propertySets.Item("ProjectIndofmormation");

                //get properties
                documentNumber = propertiesProject.Item("Document Number");
                revision = propertiesProject.Item("Revision");
                projectName = propertiesProject.Item("Project Name");

                //set properties

                documentNumber.set_Value(1); //set value requires a string, TEST if int is valid
                //
                //


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
    }
}