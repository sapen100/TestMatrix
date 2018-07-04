using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO;
using System.Collections;
using Microsoft.Office.Core;
using excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;


namespace TestMatrixApp
{
    class ExcelScap : ExcelClass
    {
        ScapMethods scap;
        //Constructor
        public ExcelScap(ScapMethods scap)
        {
            this.scap = scap;
        }        

        public override void readTablesFromWord(string wordFile)
        {
            using (DocX document = DocX.Load(wordFile))
            {                
                string cvss = "";
                if (document.Tables.Count > 0)
                {
                    foreach (Novacode.Table table in document.Tables)
                    {
                        if (scap.myTable.containsAffectedSoftware(table.Rows[0]))
                        {
                            scap.myTable.setHeaderSupporteTrue();  // We are in a new Table, so we have restart the table variables
                            extractTableContent(table);
                        }
                        else
                        {
                            if (scap.myTable.tableContainsCVEs(table.Rows[0]))
                            {
                                cvss += scap.myTable.getCVS(table);
                                break;
                            }
                        }
                    }
                }                
                scap.set_Cvss(cvss);
            }

        }
       
      

        //Extracts the content of a Individual Table        
        private void extractTableContent(Novacode.Table table)
        {
            scap.myTable.setSystemOrComponent(table.Rows[0].Cells[0].Paragraphs[0].Text);            
            if (table.Rows[0].Cells.Count == 4)
            {
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    processRowFourCols(table.Rows[i]);                    
                    //separateSystemsOrComponentsInTwoIfRequired();
                }
            }
            else if (table.Rows[0].Cells.Count == 5)
            {
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    processRowFiveCols(table.Rows[i]);
                    //separateSystemsOrComponentsInTwoIfRequired();
                }
            }
        }


        

/*WE NEED TO TEST
WE NEED TO SAVE THE CONTENT OF A BULLETIN, FROM THE OLD AND THE NEW, SO WE COULD COMPARE THEM
WE NEED TO RELEASE THE EXEL SHEET IF A PROBLEM OCCURS*/

        //When the table has 4 columns, it could contain platform or components, that is the reason we have sysorComponent value
        public void processRowFourCols(Novacode.Row rowContent)
        {
            string content, bulletinsRep;

            if (rowContent.Cells.Count == 4)  //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                scap.writePlatformOrComponentAndKB(rowContent.Cells[0].Paragraphs[0].Text, rowContent.Cells[3].Paragraphs[0].Text );
                scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent.Cells[2].Paragraphs[0].Text;                
                content = scap.utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                bulletinsRep = scap.utilities.getTextFromCell(rowContent.Cells[3].Paragraphs);
                if (scap.myTable.isSystem())
                    scap.writeSystem(content, bulletinsRep); // we are doing this because there maybe more than one componenet    
                else
                    scap.writeComponent(content, bulletinsRep); 

            }
            
            if (rowContent.Cells.Count == 1)
            {
                content = rowContent.Cells[0].Paragraphs[0].Text;
                scap.myTable.setSupportedHeader(content);
                //myTable.setSystemOrComponent(content);
            }
        }

        //This method processes extracts information from a 5 column row
        public void processRowFiveCols(Novacode.Row rowContent)
        {
            string content, bulletinsRep, contentFirstCell;

            if ((rowContent.Cells.Count == 5) && (scap.myTable.isSystem())) //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                scap.writePlatformOrComponentAndKB(rowContent.Cells[0].Paragraphs[0].Text, rowContent.Cells[4].Paragraphs[0].Text);
                scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent.Cells[3].Paragraphs[0].Text;
                content                      = scap.utilities.getTextFromCell(rowContent.Cells[1].Paragraphs);
                bulletinsRep                 = scap.utilities.getTextFromCell(rowContent.Cells[4].Paragraphs);
                scap.writeComponent(content, bulletinsRep); // we are doing this because there maybe more than one componenet                                                 
            }

            //We are assuming that if it has 5 columns, they must have KBs, and the coponents could be in the 1st column or the 2nd column
            else
            {
                if ((rowContent.Cells.Count == 5) && (scap.myTable.isComponent())) //if it has 5 rows and it is a component, you have to check the where the component is, becouse it could be in the 1st and 2nd column
                {
                    content      = scap.utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                    bulletinsRep = scap.utilities.getTextFromCell(rowContent.Cells[4].Paragraphs);

                    if (scap.utilities.thisCellcontainsConponents(content))  //Is the information we are interested in at Colum 0?
                    {
                        scap.writePlatformOrComponentAndKB(content, bulletinsRep);
                        scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent.Cells[3].Paragraphs[0].Text;
                        scap.writeComponentsOnly(content, bulletinsRep); // we are doing this because there maybe more than one component        
                    }
                    else
                    {
                        content = scap.utilities.getTextFromCell(rowContent.Cells[1].Paragraphs);
                        contentFirstCell = scap.utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                        //comments = utilities.extractSigns(utilities.getTextFromCell(rowContent.Cells[0].Paragraphs)); //This is because maybe the comments are in the other column

                        scap.writePlatformOrComponentAndKB(content, bulletinsRep);
                        scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent.Cells[3].Paragraphs[0].Text;
                        scap.writeComponentsOnly(content, bulletinsRep, contentFirstCell); // we are doing this because there maybe more than one componenet                                                 
                    }
                }

                if (rowContent.Cells.Count == 1) {  //there is a row with only one column, so we must to check whether is a CORE install
                    content = rowContent.Cells[0].Paragraphs[0].Text;
                    scap.myTable.setSupportedHeader(content);
                    //myTable.setSystemOrComponent(content);
                }
            }
        }


        /*//This method inserts operating systems to the TM if the application is described in Applications.xlsx
        private void insertSystems()
        {
            string platform = scap.getFromSheet(scap.getRow() - 1, scap.getCdAt("platform"));  
            if (platform.Equals("")) //This means that if it was an application the one that was inserted, we gotta see if there are platforms to insert for this application(e.g explorer 7.0)
            {
                
                string component = scap.getFromSheet(scap.getRow() - 1, scap.getCdAt("component"));               
                component = scap.utilities.removeContentIn(component, "()");
                ArrayList list = scap.applications.getSystems(component);
                if (list != null)
                {
                    for (int i = 0; i < list.Count; i = i + 2)
                    {
                        if (i == 0)
                            scap.insertSystemSameLine(component, (string)list[i], (string)list[i + 1]);
                        else
                            scap.insertSystem(component, (string)list[i], (string)list[i + 1]);
                    }
                }
            }
        }    */


        //Destructor
        ~ExcelScap()
        {           
        }

        public override void newDocument()
        {
            scap.newDocument();
        }

        public override void documentEnd()
        {
            scap.documentEnd();
        }    

        public override string getExcelResult()
        {
            return scap.getExcelResult();
        }

        public override void saveDocument()
        {
            scap.saveDocument();
        }

        public override void set_Bulletin(string input)
        {
            scap.set_Bulletin(input);
        }

        public override void set_Description(string input)
        {
            scap.set_Description(input);
        }

        public override void set_Kb(string input)
        {
            scap.set_Kb(input);
        }
    }
}
