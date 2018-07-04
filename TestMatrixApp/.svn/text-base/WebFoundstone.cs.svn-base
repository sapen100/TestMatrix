using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO;
using System.Collections;
using Microsoft.Office.Core;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using excel = Microsoft.Office.Interop.Excel;


namespace TestMatrixApp
{
    class WebFoundstone : ExcelClass
    {
        FoundstoneMethods foundstone;
        //Constructor
        public WebFoundstone(FoundstoneMethods foundstone)
        {
            this.foundstone = foundstone;
        }
      


        public override void readTablesFromWeb(HtmlDocument page)
        {            
            string cvss = "";
            string content;
            bool firstCVETable = true;
            if (page.DocumentNode != null)
            {
                if (page.DocumentNode.SelectNodes("//table[@class='dataTable']") != null)
                    foreach (HtmlNode table in page.DocumentNode.SelectNodes("//table[@class='dataTable']"))
                    {                        
                        if (table.SelectNodes("thead") != null)
                        {
                            HtmlNode head       = table.SelectSingleNode("thead");
                            if (head.SelectNodes("tr") != null)
                            {
                                HtmlNode headRow            = head.SelectSingleNode("tr");
                                HtmlNodeCollection headCols = headRow.SelectNodes("th");
                                if ((headCols.Count == 4) || (headCols.Count == 5) && (foundstone.myTable.isValidHeader(headCols)))//This table have 4 or 5 colums, so it is a valid table
                                {
                                    foundstone.myTable.setHeaderSupporteTrue();
                                    foundstone.myTable.setSystemOrComponent(headCols[0].InnerText);
                                    if (table.SelectNodes("tbody") != null)
                                    {
                                        foreach (HtmlNode body in table.SelectNodes("tbody"))
                                        {
                                            if (body.SelectNodes("tr") != null)
                                            {
                                                HtmlNode bodyRows = body.SelectSingleNode("tr");
                                                if ((bodyRows.SelectNodes("th") != null))
                                                {
                                                    HtmlNode bodyHead = bodyRows.SelectSingleNode("th");  // we are assuming that there is gonna be only on <th> tag
                                                    content = bodyHead.InnerText;
                                                    foundstone.myTable.setSupportedHeader(content);
                                                }
                                                else
                                                {
                                                    foreach (HtmlNode row in body.SelectNodes("tr"))
                                                    {
                                                        if (row.SelectNodes("td") != null)
                                                        {
                                                            HtmlNodeCollection cols = row.SelectNodes("td");
                                                            if (cols.Count == 4)   // This are the only rows that we are interested in 
                                                                processRowFourCols(cols);

                                                            if (cols.Count == 5)
                                                                processRowFiveCols(cols);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (table.SelectNodes("caption") != null)
                                {
                                    HtmlNode caption = table.SelectSingleNode("caption");
                                    if (foundstone.myTable.isCVEHeader(caption.InnerText))
                                    {
                                        if (firstCVETable)
                                        {                                            
                                            writePatchOnly();
                                            firstCVETable = false;
                                        }
                                        foundstone.inspectVulneravilityTable(table);
                                        //cvss += foundstone.myTable.getCVS(headCols);
                                        break;
                                    }
                                }
                            }
                        }                        
                    }
                    foundstone.set_Cvss(cvss);
                }
            if (firstCVETable)//We could not found a CVE Table
            {
                writePatchOnly();
            }

        }           

        private void writePatchOnly()
        {
            foundstone.patchEnd = foundstone.row - 1;
            foundstone.writeCVEandDescription("Patch Only", "", foundstone.patchIni);
        }

        //When the table has 4 columns, it could contain platform or components, that is the reason we have sysorComponent value
        public void processRowFourCols(HtmlNodeCollection rowContent)
        {
            /*string content, bulletinsRep;
            scap.writePlatformOrComponentAndKB(rowContent[0].InnerText, rowContent[3].InnerText);
            scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent[2].InnerText;
            content = rowContent[0].InnerText;
            bulletinsRep = rowContent[3].InnerText;
            if (scap.myTable.isSystem())
                scap.writeSystem(content, bulletinsRep); // we are doing this because there maybe more than one componenet    
            else
                scap.writeComponent(content, bulletinsRep);*/
            string content;            
            foundstone.writePlatformOrComponentAndKB(rowContent[0].InnerText);
            content = rowContent[0].InnerText;
            if (foundstone.myTable.isSystem())
                foundstone.writeSystem(content); // we are doing this because there maybe more than one componenet    
            else
                foundstone.writeComponent(content); // we are doing this because there maybe more than one componenet                                           
        }

        //This method processes extracts information from a 5 column row
        public void processRowFiveCols(HtmlNodeCollection rowContent)
        {
            string content;
            string contentFirstCell;
            if (foundstone.myTable.isSystem()) //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                foundstone.writePlatformOrComponentAndKB(rowContent[0].InnerText);
                content = rowContent[1].InnerText;
                foundstone.writeComponent(content); // we are doing this because there maybe more than one componenet                                                 
            }

            //We are assuming that if it has 5 columns, they must have BKs, and the coponents could be in the 1st column or the 2nd column
            if (!(foundstone.myTable.isSystem())) //if it has 5 rows and it is a component, you have to check the where the component is, becouse it could be in the 1st and 2nd column
            {
                if (foundstone.thisCellcontainsConponents(rowContent[0].InnerText))
                {
                    content = rowContent[0].InnerText;
                    foundstone.writePlatformOrComponentAndKB(content);
                    foundstone.writeComponentsOnly(content); // we are doing this because there maybe more than one componenet        ESTA LINEA LA ESTA CAGANDO
                }
                else
                {
                    content = rowContent[1].InnerText;
                    contentFirstCell = rowContent[0].InnerText;
                    foundstone.writePlatformOrComponentAndKB(content);
                    foundstone.writeComponentsOnly(content, contentFirstCell); // we are doing this because there maybe more than one componenet                                                 
                }
            }
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
}
