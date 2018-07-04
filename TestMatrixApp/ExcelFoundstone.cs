using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Text.RegularExpressions;
using Novacode;
using Microsoft.Office.Core;

namespace TestMatrixApp
{
    class ExcelFoundstone : ExcelClass
    {
        FoundstoneMethods foundstone;        
        //Constructor
        public ExcelFoundstone(FoundstoneMethods foundstone){            
            this.foundstone = foundstone;
        }                            

        public override void readTablesFromWord(string wordFile)
        {
            bool flag            = true;            

            using (DocX document = DocX.Load(wordFile))
            {                
                if (document.Tables.Count > 0)
                {
                    foreach (Novacode.Table table in document.Tables)
                    {
                        if ( (foundstone.myTable.containsAffectedSoftware(table.Rows[0])) )
                        {
                            extractTableContent(table);                         
                        }
                        else
                        {
                            if (foundstone.myTable.tableContainsCVEs(table.Rows[0])) 
                            {
                                if (flag)   //only the first time, because there migth be more than two tables containing CVEs
                                {
                                    writePatchIni();
                                    flag = false;
                                }
                                foundstone.inspectVulneravilityTable(table);
                            }                            
                        }
                    }
                }                
            }
            if (flag)// we could not find a cveTable
            {
                writePatchIni();
            }

        }

        private void writePatchIni()
        {
            foundstone.patchEnd = foundstone.row -1;
            foundstone.writeCVEandDescription("Patch Only", "", foundstone.patchIni);
        }

        //Extracts the content of a Individual Table        
        private void extractTableContent(Novacode.Table table){
            foundstone.myTable.setSystemOrComponent(table.Rows[0].Cells[0].Paragraphs[0].Text);            
            //myTable.setSystemOrComponent(table.Rows[0]);            
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
                //string ExecutionTimeTaken;        
                for (int i = 1; i < table.Rows.Count; i++)
                {                    
                    processRowFiveCols(table.Rows[i]);                 
                }
            }
        }

        //When the table has 4 columns, it could contain platform or components, that is the reason we have sysorComponent value
        public void processRowFourCols(Novacode.Row rowContent)
        {
            string content;
            if (rowContent.Cells.Count == 4)  //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                foundstone.writePlatformOrComponentAndKB(rowContent.Cells[0].Paragraphs[0].Text);
                content = foundstone.utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                if (foundstone.myTable.isSystem())
                    foundstone.writeSystem(content); // we are doing this because there maybe more than one componenet    
                else
                    foundstone.writeComponent(content); // we are doing this because there maybe more than one componenet                   
            }
        }


        //This method processes extracts information from a 5 column row
        public void processRowFiveCols(Novacode.Row rowContent)
        {
            string content;
            string contentFirstCell;
            if ((rowContent.Cells.Count == 5) && (foundstone.myTable.isSystem())) //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                foundstone.writePlatformOrComponentAndKB(rowContent.Cells[0].Paragraphs[0].Text);
                content = foundstone.utilities.getTextFromCell(rowContent.Cells[1].Paragraphs);
                foundstone.writeComponent(content); // we are doing this because there maybe more than one componenet                                                 
            }

            //We are assuming that if it has 5 columns, they must have BKs, and the coponents could be in the 1st column or the 2nd column
            if ((rowContent.Cells.Count == 5) && !(foundstone.myTable.isSystem())) //if it has 5 rows and it is a component, you have to check the where the component is, becouse it could be in the 1st and 2nd column
            {
                if (foundstone.thisCellcontainsConponents(rowContent.Cells[0].Paragraphs[0].Text))
                {
                    content = foundstone.utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                    foundstone.writePlatformOrComponentAndKB(content);
                    foundstone.writeComponentsOnly(content); // we are doing this because there maybe more than one componenet        ESTA LINEA LA ESTA CAGANDO
                }
                else
                {
                    content = foundstone.utilities.getTextFromCell(rowContent.Cells[1].Paragraphs);
                    contentFirstCell = foundstone.utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                    foundstone.writePlatformOrComponentAndKB(content);
                    foundstone.writeComponentsOnly(content, contentFirstCell); // we are doing this because there maybe more than one componenet                                                 
                }

            }
        }            

        //Destructor
        ~ExcelFoundstone() {         
        }

        //New document specifies that we will traverse a new document, this could be a new webpage or a it could be a new word document
        public override void newDocument()
        {
            foundstone.newDocument();
        }

        public override void documentEnd()
        {
            foundstone.documentEnd();
        }

        public override string getExcelResult()
        {
            return foundstone.getExcelResult();
        }

        public override void saveDocument()
        {
            foundstone.saveDocument();
        }

        public override void set_Bulletin(string input)
        {
            foundstone.set_Bulletin(input);
        }

        public override void set_Description(string input)
        {
            foundstone.set_Description(input);
        }

        public override void set_Kb(string input)
        {
            foundstone.set_Kb(input);
        }   
    }



        //This class was created to accelarate the data retrival from the PATCH ONLY section 
        //We save all the patch only section(w/o platforms) and when we reach the CVE table, we go come to this class to ask for the data that is needed

        class PatchOnlyCells
        {    
            LinkedList<string[]> content;
            Utilities utilities;
            public PatchOnlyCells(Utilities utilities)
            {
                this.utilities = utilities;
                content = new LinkedList<string[]>();
            }

            public void add(string platform, string application, string kb)
            {
                content.AddFirst(new string[] { platform, application, kb });
            }

            public string[] getContent(string platform, string application)
            {
                platform = platform.Trim();
                application = application.Trim();
                foreach (string[] x in content)
                {                
                    if ((x[0].Equals(platform)) && (x[1].Equals(application)))
                        return x;
                }
                return null;
            }


            public string[] getContentBasedOnComponent(string application)
            {            
                application = application.Trim();            
                foreach (string[] x in content)
                {                
                    if (x[1].Equals(application))
                        return x;
                    if (utilities.removeContentIn( x[1], "()").Equals(application))
                        return x;
                }            

                return null;
            }

            public void show() {                        
                foreach (string[] x in content) { 
                    System.Windows.Forms.MessageBox.Show("platform=" +  x[0] +" component=" +x[1]+" kb=" +x[2]);
                }            
            }
        } 
}
