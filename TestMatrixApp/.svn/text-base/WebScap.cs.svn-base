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
    class WebScap : ExcelClass
    {
        ScapMethods scap;
        //Constructor
        public WebScap(ScapMethods scap)
        {
            this.scap = scap;
        }
      


        public override void readTablesFromWeb(HtmlDocument page)
        {            
            string cvss = "";
            string content;            
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
                                if ((headCols.Count == 4) || (headCols.Count == 5) && (scap.myTable.isValidHeader(headCols)))//This table have 4 or 5 colums, so it is a valid table
                                {
                                    scap.myTable.setHeaderSupporteTrue();
                                    scap.myTable.setSystemOrComponent(headCols[0].InnerText);
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
                                                    scap.myTable.setSupportedHeader(content);
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
                                    if (scap.myTable.isCVEHeader(caption.InnerText))
                                    {
                                        cvss += scap.myTable.getCVS(headCols);
                                        break;
                                    }
                                }
                            }
                        }                        
                    }
                    scap.set_Cvss(cvss);
                }

            /*if (document.Tables.Count > 0)
            {
                foreach (Novacode.Table table in document.Tables)
                {
                    if (myTable.containsAffectedSoftware(table.Rows[0]))
                    {
                        extractTableContent(table);
                    }
                    else
                    {
                        if (myTable.tableContainsCVEs(table.Rows[0]))
                        {
                            cvss += myTable.getCVS(table);
                            break;
                        }
                    }
                }
            }
            set_Cvss(cvss);            */
        }           



        //When the table has 4 columns, it could contain platform or components, that is the reason we have sysorComponent value
        public void processRowFourCols(HtmlNodeCollection rowContent)
        {
            string content, bulletinsRep;
            scap.writePlatformOrComponentAndKB(rowContent[0].InnerText, rowContent[3].InnerText);
            scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent[2].InnerText;
            content = rowContent[0].InnerText;
            bulletinsRep = rowContent[3].InnerText;
            if (scap.myTable.isSystem())
                scap.writeSystem(content, bulletinsRep); // we are doing this because there maybe more than one componenet    
            else
                scap.writeComponent(content, bulletinsRep);
            
        }

        //This method processes extracts information from a 5 column row
        public void processRowFiveCols(HtmlNodeCollection rowContent)
        {
            string content, bulletinsRep, contentFirstCell;

            if (scap.myTable.isSystem()) //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                scap.writePlatformOrComponentAndKB(rowContent[0].InnerText, rowContent[4].InnerText);
                scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent[3].InnerText;
                content = rowContent[1].InnerText;
                bulletinsRep = rowContent[4].InnerText;
                scap.writeComponent(content, bulletinsRep); // we are doing this because there maybe more than one componenet                                                 
            }
            //We are assuming that if it has 5 columns, they must have KBs, and the coponents could be in the 1st column or the 2nd column
            else
            {
                if (scap.myTable.isComponent()) //if it has 5 rows and it is a component, you have to check the where the component is, becouse it could be in the 1st and 2nd column
                {
                    content = rowContent[0].InnerText;
                    bulletinsRep = rowContent[4].InnerText;

                    if (scap.utilities.thisCellcontainsConponents(content))  //Is the information we are interested in at Colum 0?
                    {
                        scap.writePlatformOrComponentAndKB(content, bulletinsRep);
                        scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent[3].InnerText;
                        scap.writeComponentsOnly(content, bulletinsRep); // we are doing this because there maybe more than one component        
                    }
                    else
                    {
                        content = rowContent[1].InnerText;
                        contentFirstCell = rowContent[0].InnerText;
                        //comments = utilities.extractSigns(utilities.getTextFromCell(rowContent.Cells[0].Paragraphs)); //This is because maybe the comments are in the other column

                        scap.writePlatformOrComponentAndKB(content, bulletinsRep);
                        scap.sheet.Cells[scap.getRow(), scap.getCdAt("risk")] = rowContent[3].InnerText;
                        scap.writeComponentsOnly(content, bulletinsRep, contentFirstCell); // we are doing this because there maybe more than one componenet                                                 
                    }
                }                
            }
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
