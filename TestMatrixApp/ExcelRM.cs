//NOTE: ExcelRM is exactly the same as ExcelScap.cs, with the only difference that RM does not support applications Office 64 bits
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
    class ExcelRM : ExcelClass
    {
        //Global Variables
        Dictionary<string, int> cd    = new Dictionary<string, int>() { { "index", 1 }, { "bulletin", 2 }, { "cve", 3 }, { "description", 4 }, { "globalKb", 5 }, { "kb", 6 }, { "risk", 7 }, { "platform", 8 }, { "component", 9 }, { "status", 10 }, { "na", 13 }, { "supported", 14 } };
        const string originalTemplate = "OriginalRMTM.xlsx";
        string excelResult            = "TMResult.xlsx";                        
        int row, col, index, tableIniRow;        

        Table myTable;
        Applications applications;
        Utilities utilities;
        ExcelUtilities excelUtils;
        

        excel.Application application;
        excel.Workbooks workBooks;
        excel.Workbook workBook;
        excel.Sheets sheets;
        excel.Worksheet sheet;


        string bulletin, cvss, description, kb;

        //Constructor
        public ExcelRM(string applicationLocation, string configPath, ErrorMessages em, Utilities utilities)
        {
            this.utilities  = utilities;
            excelResult     = applicationLocation + "\\" + excelResult;
            myTable         = new Table(cd["platform"], cd["component"], configPath, em);
            applications    = new Applications(utilities, configPath, "RMApplications");            
            excelUtils      = new ExcelUtilities(applications, utilities, cd);
            row             = 3;   //RM template starts at this row            
            tableIniRow     = row;
            index           = 0;        
        }

        private void openExcelTemplate()
        {
            createExcelResult();
            if (File.Exists(excelResult))
            {
                application = new excel.Application();
                application.DisplayAlerts = false;
                workBooks = application.Workbooks;
                workBook = workBooks.Open(excelResult, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheets = workBook.Sheets;
                sheet = sheets.get_Item(1);
                excelUtils.setSheet(sheet);
            }
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
                set_Cvss(cvss);
            }

        }


        //Copies the original template to a TM result 
        private void createExcelResult()
        {
            try
            {
                if (File.Exists(excelResult))
                    File.Delete(excelResult);
                File.Copy(originalTemplate, excelResult);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("ERROR: " + ex.Message);                
                System.Environment.Exit(0);
            }
        }

        //It processes a new word Document
        public override void newDocument()
        {
            if (index == 0)
                openExcelTemplate();
            index++;
            tableIniRow = row;
        }

        //Extracts the content of a Individual Table        
        private void extractTableContent(Novacode.Table table)
        {
            //myTable.setSystemOrComponent(table.Rows[0]);
            myTable.setSystemOrComponent(table.Rows[0].Cells[0].Paragraphs[0].Text);            
            if (table.Rows[0].Cells.Count == 4)
            {
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    processRowFourCols(table.Rows[i]);                                        
                }
            }
            else if (table.Rows[0].Cells.Count == 5)
            {
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    processRowFiveCols(table.Rows[i]);                    
                }
            }
        }

        //This method inserts operating systems to the TM if the application is described in Applications.xlsx
        private void insertSystems()
        {
            string platform = (sheet.Cells[row - 1, cd["platform"]].Text).Trim();            
            if (platform.Equals("")) //This means that if it was an application the one that was inserted, we gotta see if there are platforms to insert for this application(e.g explorer 7.0)
            {
                string component = (sheet.Cells[row - 1, cd["component"]].Text).Trim();
                ArrayList list = applications.getSystems(component);
                if (list != null)
                {
                    for (int i = 0; i < list.Count; i = i + 2)
                    {
                        if (i == 0)
                            insertSystemSameLine(component, (string)list[i], (string)list[i + 1]);
                        else
                            insertSystem(component, (string)list[i], (string)list[i + 1]);
                    }
                }
            }
        }

        //This method is a insertSystems method helper that inserts a system if needed
        private void insertSystem(string component, string system, string supported)
        {
            sheet.Cells[row, cd["kb"]]        = sheet.Cells[row - 1, cd["kb"]].Text;
            sheet.Cells[row, cd["risk"]]      = sheet.Cells[row - 1, cd["risk"]].Text;
            sheet.Cells[row, cd["platform"]]  = system;
            sheet.Cells[row, cd["component"]] = component;            
            row++;

        }

        //This method is a insertSystems method helper that inserts a system if needed.
        private void insertSystemSameLine(string component, string system, string supported)
        {
            sheet.Cells[row - 1, cd["platform"]]  = system;            
        }

        //When the table has 4 columns, it could contain platform or components, that is the reason we have sysorComponent value
        public void processRowFourCols(Novacode.Row rowContent)
        {
            string content;
            if (rowContent.Cells.Count == 4)   //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                writePlatformOrComponentAndKB(rowContent.Cells[0].Paragraphs[0].Text);                
                sheet.Cells[row, cd["risk"]] = rowContent.Cells[2].Paragraphs[0].Text;                
                content = utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                if (myTable.isSystem())
                    writeSystem(content); // we are doing this because there maybe more than one componenet    
                else
                    writeComponent(content); // we are doing this because there maybe more than one componenet   
            }
        }

        //This method processes extracts information from a 5 column row
        public void processRowFiveCols(Novacode.Row rowContent)
        {
            string content, contentFirstCell;

            if ((rowContent.Cells.Count == 5) && (myTable.isSystem())) //there are some cases in which the tool finds a cell in the middle that has a comment and NOT actual data
            {
                //content = utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                writePlatformOrComponentAndKB(rowContent.Cells[0].Paragraphs[0].Text);
                sheet.Cells[row, cd["risk"]] = rowContent.Cells[3].Paragraphs[0].Text;
                content = utilities.getTextFromCell(rowContent.Cells[1].Paragraphs);
                writeComponent(content); // we are doing this because there maybe more than one componenet                                                 
            }

            //We know it has 5 columns, they must have BKs, and the coponents could be in the 1st column or the 2nd column
            else 
                if ((rowContent.Cells.Count == 5) && !(myTable.isSystem())) //if it has 5 rows and it is a component, you have to check the where the component is, becouse it could be in the 1st and 2nd column
                    {                        
                        content  = utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);                            

                        if ( utilities.thisCellcontainsConponents(content) )
                        {                            
                            writePlatformOrComponentAndKB(content);
                            sheet.Cells[row, cd["risk"]] = rowContent.Cells[3].Paragraphs[0].Text;
                            writeComponentsOnly(content); // we are doing this because there maybe more than one component        
                        }
                        else
                        {                            
                            content          = utilities.getTextFromCell(rowContent.Cells[1].Paragraphs);
                            contentFirstCell = utilities.getTextFromCell(rowContent.Cells[0].Paragraphs);
                            //comments = utilities.extractSigns(utilities.getTextFromCell(rowContent.Cells[0].Paragraphs)); //This is because maybe the comments are in the other column
                            
                            writePlatformOrComponentAndKB(content);                            
                            sheet.Cells[row, cd["risk"]] = rowContent.Cells[3].Paragraphs[0].Text;
                            writeComponentsOnly(content, contentFirstCell); // we are doing this because there maybe more than one componenet                                                 
                        }
                    }
        }


        /*private void writeSuported() 
        {
            if (myTable.isSystem())
            {
                if ((sheet.Cells[row, cd["platform"]].Text).ToUpper().Contains("ITANIUM"))
                    sheet.Cells[row, cd["supported"]] = "No";
                else
                    sheet.Cells[row, cd["supported"]] = "Yes";
            }
        }*/
       
        //This method is trigered whenever the table we are working is has 5 columns and it is a SYSTEM TABLE , that means that the only column that many contain mone than one content is the COMPONENT column
        private void writeComponent(string content)
        {
            string comments     = "";
            string[] components = utilities.deleteCommentsInTheMiddle( content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries) );
            if (components.Length == 1)
            {
                if (!(components[0].Trim().ToUpper().Equals("NOT APPLICABLE")))
                {
                    comments = utilities.extractComments(components[0]);
                    sheet.Cells[row, cd["component"]] = utilities.deleteComments(components[0]);
                    sheet.Cells[row, cd["kb"]] = sheet.Cells[row, cd["kb"]].Text + comments;
                }
                row++;
                separateSystemsOrComponentsInTwoIfRequired();
            }
            else
            {
                try
                {
                    int i;
                    string risk = sheet.Cells[row, cd["risk"]].Text;
                    string platform = sheet.Cells[row, cd["platform"]].Text;

                    for (i = 0; i < components.Length; i += 2)
                    {
                        comments = utilities.extractComments(components[i]);
                        sheet.Cells[row, cd["component"]] = utilities.deleteComments(components[i]);
                        sheet.Cells[row, cd["risk"]] = risk;
                        sheet.Cells[row, cd["platform"]] = platform;
                        sheet.Cells[row, cd["kb"]] = utilities.deleteParenthesis(components[i + 1]).Trim() + comments;                        
                        row++;
                        separateSystemsOrComponentsInTwoIfRequired();
                    }
                }
                catch (Exception e){
                    System.Windows.Forms.MessageBox.Show("Error Writing Components: " + e.Message  + "\n\n Please verify if the following data is consistent: " + getComponents(components));
                } 
            }
        }

        //This method is trigered whenever the table we are working is a SYSTEM TABLE , that means that the only column that many contain mone than one content is the system column
        private void writeSystem(string content)
        {
            string comments = "";
            string[] systems = content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (systems.Length == 1)
            {
                if (!(systems[0].Trim().ToUpper().Equals("NOT APPLICABLE")))
                {
                    comments = utilities.extractComments(systems[0]);
                    sheet.Cells[row, cd["platform"]] = utilities.deleteComments(systems[0]);
                    sheet.Cells[row, cd["kb"]] = sheet.Cells[row, cd["kb"]].Text + comments;
                }
                row++;
                separateSystemsOrComponentsInTwoIfRequired();
            }
            else
            {
                try
                {
                    int i;
                    string risk = sheet.Cells[row, cd["risk"]].Text;
                    string component = sheet.Cells[row, cd["component"]].Text;

                    for (i = 0; i < systems.Length; i += 2)
                    {
                        comments = utilities.extractComments(systems[i]);
                        sheet.Cells[row, cd["platform"]] = utilities.deleteComments(systems[i]);
                        sheet.Cells[row, cd["risk"]] = risk;
                        sheet.Cells[row, cd["component"]] = component;
                        sheet.Cells[row, cd["kb"]] = utilities.deleteParenthesis(systems[i + 1]).Trim() + comments;
                        row++;
                        separateSystemsOrComponentsInTwoIfRequired();
                    }
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Error Writing Components: " + e.Message + "\n\n Please verify if the following data is consistent: " + getComponents(systems));
                }
            }
        }

        //This method is trigered whenever the content has more than one item in it, so it needs to be splitted.
        // @_otherCellContent is used to extract comments, that may be in the other column, and also to fix some Microsoft errors. (the content is not necesarily the one where the KB is located. it could also be in the other cell, that is what @othercellcontent contains)
        private void writeComponentsOnly(string content, string otherCellContent = "")
        {
            string comments           = "";
            string cleanContent       = "";
            string otherCellComments  = utilities.extractComments(otherCellContent);
                   otherCellContent   = utilities.deleteComments(otherCellContent);
            
            string[] components = utilities.deleteCommentsInTheMiddle( content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries) );

            if (components.Length == 1) //is there more than one component in the cell
            {
                cleanContent = utilities.deleteComments(components[0]);
                sheet.Cells[row, cd["component"]] = cleanContent;
                row++;                
                excelUtils.replaceComponentContentIfRequired(cleanContent, otherCellContent, row);
                separateSystemsOrComponentsInTwoIfRequired();
            }
            else
            {
                int i;
                string risk = sheet.Cells[row, cd["risk"]].Text;

                for (i = 0; i < components.Length; i += 2)
                {
                    comments = utilities.extractComments(components[i].Trim()) + otherCellComments;
                    cleanContent = utilities.deleteComments(components[i]);
                    sheet.Cells[row, cd["component"]] = cleanContent;
                    sheet.Cells[row, cd["risk"]]      = risk;
                    sheet.Cells[row, cd["kb"]]        = utilities.deleteParenthesis(components[i + 1]) + comments;
                    row++;
                    excelUtils.replaceComponentContentIfRequired(cleanContent, otherCellContent, row);
                    separateSystemsOrComponentsInTwoIfRequired();
                }
            }
        }
        

        private string getComponents(string[] components)
        {
            string result = "\n";
            foreach (string component in components) 
            {
                result = result + component + "\n";
            }
            return result;
        }

        //This method ducplicates the last row, whenever that cell contains a cell containing more than one content. For example, cell contains "windows xp and windows 2003"
        private void separateSystemsOrComponentsInTwoIfRequired()
        {
            if (utilities.moreThanOneContent(sheet.Cells[row - 1, cd["platform"]].Text))
            {
                string strSystems = sheet.Cells[row - 1, cd["platform"]].Text;
                string[] systems = utilities.separateAndOccurrences(strSystems);                    
                sheet.Cells[row - 1, cd["platform"]] = systems[0];
                sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]];
                sheet.Cells[row, cd["risk"]] = sheet.Cells[row - 1, cd["risk"]];
                sheet.Cells[row, cd["component"]] = sheet.Cells[row - 1, cd["component"]];                
                sheet.Cells[row, cd["platform"]] = systems[1];                
                row++;
            }

            else if (utilities.moreThanOneContent(sheet.Cells[row - 1, cd["component"]].Text))
            {                
                string strComponents = sheet.Cells[row - 1, cd["component"]].Text;
                string[] components  = utilities.separateAndOccurrences(strComponents);                    
                sheet.Cells[row - 1, cd["component"]] = components[0].Trim() + utilities.getBitEdition(sheet.Cells[row - 1, cd["component"]].Text);

                string tmpPlatform      = sheet.Cells[row - 1, cd["platform"]].Text; //Before it starts putting more platforms, we copy what is the original
                
                insertSystems();

                sheet.Cells[row, cd["kb"]]        = sheet.Cells[row - 1, cd["kb"]];
                sheet.Cells[row, cd["risk"]]      = sheet.Cells[row - 1, cd["risk"]];                
                sheet.Cells[row, cd["platform"]]  = tmpPlatform;
                sheet.Cells[row, cd["component"]] = components[1].Trim();                
                row++;
            }
            insertSystems(); 
        }

      

        //Once we completed traversing a word Document, we write the text that we got from the word and merge the cells that need to be merged
        public override void documentEnd()
        {
            if (tableIniRow != row) //nothing was add, so there is nothing to merge
            {
                sheet.Cells[tableIniRow, cd["index"]] = get_Index().ToString();
                sheet.Cells[tableIniRow, cd["bulletin"]] = get_Bulletin();
                sheet.Cells[tableIniRow, cd["cve"]] = get_Cvss();
                sheet.Cells[tableIniRow, cd["description"]] = get_Description();
                sheet.Cells[tableIniRow, cd["globalKb"]] = get_Kb();
                mergeCells();
                tableIniRow = row;
            }
        }


        //Save the excel created
        public override void saveDocument()
        {            
            workBook.Save();
            workBook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            application.Quit();            
        }

        //Merge cells
        private void mergeCells()
        {
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["index"]], sheet.Rows.Cells[row - 1, cd["index"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["bulletin"]], sheet.Rows.Cells[row - 1, cd["bulletin"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["cve"]], sheet.Rows.Cells[row - 1, cd["cve"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["description"]], sheet.Rows.Cells[row - 1, cd["description"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["globalKb"]], sheet.Rows.Cells[row - 1, cd["globalKb"]]].Merge();
        }

        //This method writes the KB and the platform or component.
        private void writePlatformOrComponentAndKB(string content)
        {
            int endPos;
            string comments      = utilities.extractComments(content);
            string localPlatform = "";
            string localKb       = "";
            string tempContent   = content;

            int iniPos = content.IndexOf("(KB");            
            if (iniPos != -1)
            {
                localPlatform = content.Substring(0, iniPos);
                content       = content.Substring(iniPos);
                endPos        = content.IndexOf(")");
                if (endPos != -1)
                {
                    localKb                                          = content.Substring(1, endPos - 1);
                    sheet.Cells[row, myTable.getSystemOrComponent()] = utilities.deleteComments(localPlatform);   //table knows if the table has components or OS in the first row
                    sheet.Cells[row, cd["kb"]]                       = localKb + comments;
                    return;
                }
            }            
            sheet.Cells[row, cd["kb"]] = get_Kb() + comments;
            sheet.Cells[row, myTable.getSystemOrComponent()] = utilities.deleteComments(tempContent);
        }


        //Destructor
        ~ExcelRM()
        {
            releaseObject(application);
            releaseObject(workBook);
            releaseObject(workBooks);
            releaseObject(sheet);
            releaseObject(sheets);
        }

        //This method was to filter office products, but the implementation was cancelled
        private bool itIsNotOffice64Bits(string content) { 
            content = content.ToUpper();
            if ( ((content.Contains("MICROSOFT OFFICE")) || (content.Contains("MICROSOFT WORD")) ) && (content.Contains("64-bit editions")))
                return true;
            else
                return false;
        }


        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

            }
            finally
            {
                GC.Collect();
            }
        }

        public override void set_Bulletin(string input)
        {
            bulletin = input;
        }

        public void set_Cvss(string input)
        {
            cvss = input;
        }
        public override void set_Description(string input)
        {
            description = input;
        }
        public override void set_Kb(string input)
        {
            kb = input;
        }

        public override bool isResultOpen()
        {
            return (index != 0) ? true : false;
        }

        public string get_Bulletin()
        {
            return bulletin;
        }

        public string get_Cvss()
        {
            return cvss;
        }

        public string get_Description()
        {
            return description;
        }


        public string get_Kb()
        {
            return kb;
        }

        public int get_Index()
        {
            return index;
        }

        public override string getExcelResult()
        {
            return excelResult;
        }       

    }
}
