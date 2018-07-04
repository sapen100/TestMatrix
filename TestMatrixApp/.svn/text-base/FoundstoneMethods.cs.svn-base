using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Collections;
using System.Text.RegularExpressions;
using excel = Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;

namespace TestMatrixApp
{
    class FoundstoneMethods
    {
        //Global Variables
        string[] applicableStrs       = { "IMPORTANT", "LOW", "CRITICAL", "MODERATE" };
        Dictionary<string, int> cd    = new Dictionary<string, int>() { { "index", 1 }, { "bulletin", 2 }, { "globalKb", 3 }, { "cve", 4 }, { "description", 5 }, { "kb", 6 }, { "platform", 8 }, { "component", 9 } };
        string[] invalidHeaders       = { "Developer Tools" };
        string excelResult            = "TMResult.xlsx";
        const string originalTemplate = "OriginalFoundstoneTM.xlsx";
        string configPath;
        public int row;        
        int index;
        int tableIniRow;
        public int patchIni;
        public int patchEnd;

        public Table myTable;
        public Applications applications;
        public Utilities utilities;
        public ExcelUtilities excelUtils;
        Stopwatch sw;
        ErrorMessages em;

        excel.Application application;
        excel.Workbooks workBooks;
        excel.Workbook workBook;
        excel.Sheets sheets;
        excel.Worksheet sheet;
        PatchOnlyCells patchOnlyContent;


        string bulletin, cvss, description, kb;

        public FoundstoneMethods(string applicationLocation, string configPath, ErrorMessages em, Utilities utilities)
        {
            this.configPath = configPath;
            this.utilities = utilities;
            this.em = em;
            excelResult = applicationLocation + "\\" + excelResult;
            myTable = new Table(cd["platform"], cd["component"], configPath, em);
            applications = new Applications(utilities, configPath, "FoundstoneApplications");
            excelUtils = new ExcelUtilities(applications, utilities, cd);
            patchOnlyContent = new PatchOnlyCells(utilities);
            sw = new Stopwatch();
            row = 2;   //our template starts at this row            
            tableIniRow = row;
            index = 0;                       
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
            else
            {
                em.showErrorFileNotExist(excelResult);
                System.Environment.Exit(0);
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
        public void newDocument()
        {
            if (index == 0)
                openExcelTemplate();
            index++;
            tableIniRow = row;
            patchIni = row;
        }


        public void setPatchEnd(){
            patchEnd = row - 1;   //When we reach the table that contains the CVEs, we know it is the end of the patch Table(All vulnerabilities in a Bulletin)
        }


        public void writeCVEandDescription(string cve, string description, int ini)
        {            
            if (ini != row)
            {
                sheet.Cells[ini, cd["cve"]] = cve;
                sheet.Cells[ini, cd["description"]] = description;
                sheet.Rows.Range[sheet.Rows.Cells[ini, cd["cve"]], sheet.Rows.Cells[row - 1, cd["cve"]]].Merge();
                sheet.Rows.Range[sheet.Rows.Cells[ini, cd["description"]], sheet.Rows.Cells[row - 1, cd["description"]]].Merge();
            }
        }

        public void inspectVulneravilityTable(Novacode.Table table)
        {
            string headerSeparator;
            int cveIni;
            for (int c = 1; c < table.Rows[1].Cells.Count - 1; c++)   //This for traverses the number of columns that have cve
            {
                //System.Windows.Forms.MessageBox.Show("estoy inspecionando la columna " + c.ToString());
                headerSeparator = "";
                cveIni = row;
                for (int i = 2; i < table.Rows.Count; i++)      //The first row has the header of the table and the second has the the column definition
                {
                    if (table.Rows[i].Cells.Count == table.Rows[1].Cells.Count)        //It measn that this column is a header separator                    
                    {
                        if (isAffected(table.Rows[i].Cells[c].Paragraphs[0].Text))// This means that wheter the corresponding coulumn contains MODERATE, CRITICAL LOW IMPORTANT
                        {
                            findRowsToCopyFrom(headerSeparator, table.Rows[i].Cells[0].Paragraphs[0].Text);
                        }
                    }
                    else
                        headerSeparator = (table.Rows[i].Cells[0].Paragraphs[0].Text).Trim();
                }
                getCVEandDescription(table.Rows[1].Cells[c].Paragraphs[0].Text, cveIni);
            }
        }


        public void inspectVulneravilityTable(HtmlNode table)
        {            
            string headerSeparator;
            int cveIni;
            HtmlNodeCollection cols;
            HtmlNode             head = table.SelectSingleNode("thead");
            HtmlNodeCollection bodies = table.SelectNodes("tbody");

            HtmlNode headRow            = head.SelectSingleNode("tr");
            HtmlNodeCollection headCols = headRow.SelectNodes("th");            
            HtmlNodeCollection headerSeparatorH;

            //HtmlNode bodyRows = body.SelectSingleNode("tr");

            for (int c = 1; c < headCols.Count - 1; c++)   //This for traverses the number of columns that have cve
            {                
                headerSeparator = "";
                cveIni = row;
                //for (int i = 2; i < bodyRows.Count; i++)      //The first row has the header of the table and the second has the the column definition                
                foreach (HtmlNode body in bodies)
                {
                    foreach (HtmlNode HTMLRow in body.SelectNodes("tr"))      //The first row has the header of the table and the second has the the column definition
                    {

                        if ((cols = HTMLRow.SelectNodes("td")) != null)        //It means that this column is a header separator                    
                        {
                            if (isAffected(cols[c].InnerText))// This means that wheter the corresponding coulumn contains MODERATE, CRITICAL LOW IMPORTANT
                            {
                                findRowsToCopyFrom(headerSeparator, cols[0].InnerText);
                            }
                        }
                        else
                        {
                            if ((headerSeparatorH = HTMLRow.SelectNodes("th")) != null)
                                headerSeparator = (headerSeparatorH[0].InnerText).Trim();
                        }
                    }
                }
                getCVEandDescription(headCols[c].InnerText, cveIni);
            }
        }

        

        /*public void inspectVulneravilityTable(Novacode.Table table)
        {
            string headerSeparator;
            int cveIni;
            for (int c = 1; c < table.Rows[1].Cells.Count - 1; c++)   //This for traverses the number of columns that have cve
            {
                //System.Windows.Forms.MessageBox.Show("estoy inspecionando la columna " + c.ToString());
                headerSeparator = "";
                cveIni = row;
                for (int i = 2; i < table.Rows.Count; i++)      //The first row has the header of the table and the second has the the column definition
                {
                    if (table.Rows[i].Cells.Count == table.Rows[1].Cells.Count)        //It measn that this column is a header separator                    
                    {
                        if (isAffected(table.Rows[i].Cells[c].Paragraphs[0].Text))// This means that wheter the corresponding coulumn contains MODERATE, CRITICAL LOW IMPORTANT
                        {
                            findRowsToCopyFrom(headerSeparator, table.Rows[i].Cells[0].Paragraphs[0].Text);
                        }
                    }
                    else
                        headerSeparator = (table.Rows[i].Cells[0].Paragraphs[0].Text).Trim();
                }
                getCVEandDescription(table.Rows[1].Cells[c].Paragraphs[0].Text, cveIni);
            }
        }*/

        //This method is a insertSystems method helper that inserts a system if needed
        private void insertSystem(string component, string system, string suppoerted)
        {
            sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]].Text;
            //sheet.Cells[row, cd["risk"]] = sheet.Cells[row - 1, cd["risk"]].Text;
            sheet.Cells[row, cd["platform"]] = system;
            sheet.Cells[row, cd["component"]] = component;
            //sheet.Cells[row, cd["supported"]] = suppoerted;
            row++;

        }
        
        private bool isAffected(string content)
        {
            content = content.Trim().ToUpper();
            for (int i = 0; i < applicableStrs.Length; i++)
            {
                if (content.Contains(applicableStrs[i]))
                    return true;
            }
            return false;
            //return applicableStrs.Contains(content.);
        }

        private void findRowsToCopyFrom(string header, string content)
        {
            string withoutInstalledOn;
            string originalContent = utilities.deleteComments(content);
            string[] rtc;

            if ((!content.Contains(header)) || (content.Equals(header)))
                header = "";

            if (!header.Equals(""))
                content = deleteHeader(content, header);

            //string bitEdition = utilities.getBitEdition(content);
            string item;
            string[] items = utilities.separateAndOccurrences(content);
            header = utilities.deleteComments(header);
            for (int i = 0; i < items.Count(); i++)
            {
                item = utilities.deleteKBContent(utilities.deleteComments(items[i].Trim()));
                withoutInstalledOn = delIstalledOnStr(item);
                
               /* if (i + 1 == items.Count()) //Is it the last item of the array                    
                    item = utilities.deleteComments(items[i].Trim());
                else
                    item = utilities.deleteComments(items[i].Trim()) + bitEdition;*/

                if ((rtc = findRowToCopyFrom(header, item)) != null)
                {
                    copyRow(rtc);
                }
                else
                {
                    //These are all the possible wasy we could find the content in the patch only section
                    if (((rtc = findRowToCopyFrom("", item)) != null) || ((rtc = findRowToCopyFrom(item, "")) != null) || ((rtc = findRowToCopyFrom("", withoutInstalledOn)) != null) || ((rtc = findRowToCopyFrom("", originalContent)) != null) || ((rtc = findRowToCopyFrom("", utilities.deleteKBContent(originalContent))) != null) || ((rtc = findRowToCopyFromMatchComponOnly(item)) != null) || ((rtc = findRowToCopyFromMatchComponOnly(originalContent)) != null) || ((rtc = findRowToCopyFromMatchComponOnly(withoutInstalledOn)) != null) || ((rtc = findRowToCopyFrom("", utilities.removeContentIn(item, "()"))) != null) || ((rtc = findRowToCopyFrom(utilities.removeContentIn(item, "()"), "")) != null)) //after We search everything, and we could not find the RTCF, we search for the content only                                         
                        copyRow(rtc);
                    else //if we do not find the content we create a row with no KB
                        createRowWNoKb(header, item);
                }
            }
        }

        private string delIstalledOnStr(string content)
        {
            int iniPos;            

            if ((iniPos = content.ToUpper().IndexOf(" INSTALLED ON ")) != -1)
            {
                return content.Substring(iniPos + " INSTALLED ON ".Length);
            }

            if ((iniPos = content.ToUpper().IndexOf(" ON ")) != -1)
            {
                return content.Substring(iniPos + " ON ".Length);
            }
            return content;
        }

        private void getCVEandDescription(string content, int cveIni)
        {
            string cve, description;
            int iniPos = content.IndexOf("CVE-");
            if (iniPos != -1)
            {
                description = content.Substring(0, iniPos);
                cve = content.Substring(iniPos );
                writeCVEandDescription(cve, deleteExtraChars(description), cveIni);
            }
        }


        private string deleteExtraChars(string description) 
        { 
            description = description.Trim();
            if (description.EndsWith("-"))
                return description.Substring(0, description.Length - 1).Trim();
            if (description.EndsWith("–"))
                return description.Substring(0, description.Length - 1).Trim();
            if (description.EndsWith(":"))
                return description.Substring(0, description.Length - 1).Trim();
            return description;
        }


        private void createRowWNoKb(string header, string content)
        {
            if (header.Contains("Windows "))
            {
                sheet.Cells[row, cd["platform"]] = header;
                sheet.Cells[row, cd["component"]] = content;
            }
            else
                if (header.Equals(""))
                {
                    sheet.Cells[row, cd["component"]] = content;
                }
                else
                {
                    sheet.Cells[row, cd["platform"]] = content;
                    sheet.Cells[row, cd["component"]] = header;
                }
            row++;
        }

        //When we start traversing the vulneravility table, there are some headers that are not valid at all, and they are only separators
        private bool isInvalidHeader(string header)
        {
            foreach (string item in invalidHeaders)
            {
                if (item.Equals(header))
                    return true;
            }
            return false;
        }


        private string[] findRowToCopyFrom(string header, string content)
        {
            string[] result;
            if ((result = patchOnlyContent.getContent(header, content)) != null)
                return result;
            if ((result = patchOnlyContent.getContent(content, header)) != null)
                return result;
            return null;
        }



        private string[] findRowToCopyFromMatchComponOnly(string content)
        {
            string[] result;
            if ((result = patchOnlyContent.getContentBasedOnComponent(content)) != null)
                return result;            
            return null;
        }
        private void copyRow(string[] rtc)
        {
            sheet.Cells[row, cd["platform"]] = rtc[0];
            sheet.Cells[row, cd["component"]] = rtc[1];
            sheet.Cells[row, cd["kb"]] = rtc[2];
            row++;
            insertSystems();
        }

        private string deleteHeader(string content, string header)
        {
            if (content.Contains(header.Trim()))
            {
                int pos = content.IndexOf(" Windows ");
                if (pos != -1)
                    return content.Substring(pos + 1);
            }
            return content;
        }


        //This method inserts operating systems to the TM if the application loaded with the respective OS that are suppoert it. It must be described in Applications.xlsx
        private void insertSystems()
        {
            string platform = (sheet.Cells[row - 1, cd["platform"]].Text).Trim();
            //###if (!(myTable.isSystem()) && (platform.Equals(""))) //This means that if it was an application the one that was inserted, we gotta see if there are platforms to insert for this application(e.g explorer 7.0)            
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




        //This method is a insertSystems method helper that inserts a system if needed.
        private void insertSystemSameLine(string component, string system, string supported)
        {
            sheet.Cells[row - 1, cd["platform"]] = system;
            //sheet.Cells[row-1, cd["supported"]] = supported;
        }


        public bool thisCellcontainsConponents(string content)
        {
            return content.Contains("\n(KB");
        }

        //This method is trigered whenever the table we are working is has 5 columns and it is a SYSTEM TABLE , that means that the only column that many contain mone than one content is the COMPONENT column
        public void writeComponent(string content)
        {
            string[] components = utilities.splitContent(content);
            if (components.Length == 1)
            {
                if (!(components[0].Trim().ToUpper().Equals("NOT APPLICABLE")))
                {
                    sheet.Cells[row, cd["component"]] = utilities.deleteComments(components[0]);
                    sheet.Cells[row, cd["kb"]] = sheet.Cells[row, cd["kb"]].Text;
                }
                row++;
                separateSystemsOrComponentsInTwoIfRequired();
            }
            else
            {
                try
                {
                    int i;
                    string platform = sheet.Cells[row, cd["platform"]].Text;

                    for (i = 0; i < components.Length; i += 2)
                    {
                        sheet.Cells[row, cd["component"]] = utilities.deleteComments(components[i]);
                        sheet.Cells[row, cd["platform"]] = platform;
                        sheet.Cells[row, cd["kb"]] = utilities.deleteParenthesis(components[i + 1]).Trim();
                        row++;
                        separateSystemsOrComponentsInTwoIfRequired();
                    }
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Error Writing Components: " + e.Message + "\n\n Please verify the following data to be consistent: " + getComponents(components));
                }
            }
        }

        //There is more than one system in column Operatying system
        public void writeSystem(string content)
        {
            //string[] systems = content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            string[] systems = utilities.splitContent(content);
            
            if (systems.Length == 1)
            {
                if (!(systems[0].Trim().ToUpper().Equals("NOT APPLICABLE")))
                {
                    sheet.Cells[row, cd["platform"]] = utilities.deleteComments(systems[0]);
                    sheet.Cells[row, cd["kb"]] = sheet.Cells[row, cd["kb"]].Text;
                }
                row++;
                separateSystemsOrComponentsInTwoIfRequired();
            }
            else
            {
                try
                {
                    int i;
                    string component = sheet.Cells[row, cd["component"]].Text;

                    for (i = 0; i < systems.Length; i += 2)
                    {
                        sheet.Cells[row, cd["platform"]] = utilities.deleteComments(systems[i]);
                        sheet.Cells[row, cd["component"]] = component;
                        sheet.Cells[row, cd["kb"]] = utilities.deleteParenthesis(systems[i + 1]).Trim();
                        row++;
                        separateSystemsOrComponentsInTwoIfRequired();
                    }
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Error Writing Components: " + e.Message + "\n\n Please verify the following data to be consistent: " + getComponents(systems));
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


        //This method is trigered whenever the content has more than one item in it, so it needs to be splitted.
        // @_otherCellContent is used to extract comments, that may be in the other column, and also to fix some Microsoft errors. (the content is not necesarily the one where the KB is located. it could also be in the other cell, that is what @othercellcontent contains)
        public void writeComponentsOnly(string content, string otherCellContent = "")
        {
            string cleanContent = "";
            otherCellContent = utilities.deleteComments(otherCellContent);

            //string[] components = utilities.deleteCommentsInTheMiddle(content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries));
            string[] components = utilities.splitContent(content);

            if (components.Length == 1)
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
                for (i = 0; i < components.Length; i += 2)
                {
                    cleanContent = utilities.deleteComments(components[i]);
                    sheet.Cells[row, cd["component"]] = cleanContent;
                    sheet.Cells[row, cd["kb"]] = utilities.deleteParenthesis(components[i + 1]).Trim();
                    row++;
                    excelUtils.replaceComponentContentIfRequired(cleanContent, otherCellContent, row);
                    separateSystemsOrComponentsInTwoIfRequired();
                }
            }
        }

        //This method ducplicates the last row, whenever that cell contains a cell containing more than one content. For example, cell contains "windows xp and windows 2003"
        private void separateSystemsOrComponentsInTwoIfRequired()
        {
            sheet.Cells[row - 1, cd["component"]] = utilities.deleteComments(sheet.Cells[row - 1, cd["component"]].Text);
            sheet.Cells[row - 1, cd["platform"]] = utilities.deleteComments(sheet.Cells[row - 1, cd["platform"]].Text);

            if (utilities.moreThanOneContent(sheet.Cells[row - 1, cd["platform"]].Text))
            {
                string strPlatforms = sheet.Cells[row - 1, cd["platform"]].Text;
                //string[] systems = strSystems.Split(new string[] { " and " }, StringSplitOptions.RemoveEmptyEntries);
                string[] systems = utilities.separateAndOccurrences(strPlatforms);
                sheet.Cells[row - 1, cd["platform"]] = systems[0].Trim();

                addPatchOnlyCell();
                //copy Previous Cell
                sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]];
                sheet.Cells[row, cd["component"]] = sheet.Cells[row - 1, cd["component"]];
                sheet.Cells[row, cd["platform"]] = systems[1].Trim();
                row++;
                addPatchOnlyCell();
            }

            else
                if (utilities.moreThanOneContent(sheet.Cells[row - 1, cd["component"]].Text))
                {
                    string strComponents = sheet.Cells[row - 1, cd["component"]].Text;
                    string[] components = utilities.separateAndOccurrences(strComponents);
                    sheet.Cells[row - 1, cd["component"]] = components[0].Trim() + utilities.getBitEdition(sheet.Cells[row - 1, cd["component"]].Text);

                    string tmpPlatform = sheet.Cells[row - 1, cd["platform"]].Text; //Before it starts putting more platforms, we copy what is the original                

                    addPatchOnlyCell();
                    insertSystems();

                    sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]];
                    sheet.Cells[row, cd["platform"]] = tmpPlatform;
                    sheet.Cells[row, cd["component"]] = components[1].Trim();
                    row++;
                    addPatchOnlyCell();
                    insertSystems();
                }
                else
                {
                    addPatchOnlyCell();
                    insertSystems();
                }


        }

        public void documentEnd()
        {
            if (tableIniRow != row) //nothing was add, so there is nothing to merge
            {
                sheet.Cells[tableIniRow, cd["index"]] = get_Index().ToString();
                sheet.Cells[tableIniRow, cd["bulletin"]] = get_Bulletin();
                sheet.Cells[tableIniRow, cd["description"]] = get_Description();
                sheet.Cells[tableIniRow, cd["globalKb"]] = get_Kb();
                mergeCells();
                tableIniRow = row;
            }
        }

        //This method adds the last cell to path only content. It was created to increase performance because we were traversing the SpreadSheet to recover the values, so that was costly
        public void addPatchOnlyCell()
        {
            string platform = sheet.Cells[row - 1, cd["platform"]].Text;
            string component = sheet.Cells[row - 1, cd["component"]].Text;
            string kb = sheet.Cells[row - 1, cd["kb"]].Text;
            patchOnlyContent.add(platform.Trim(), component.Trim(), kb);
        }

        public void saveDocument()
        {
            if (index != 0)  // This means Result.xlsx was open, so we have to save it and close it
            {
                workBook.Save();
                workBook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
                application.Quit();
            }
        }


        //This method merges the Cells. This method is triguered 
        private void mergeCells()
        {
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["index"]], sheet.Rows.Cells[row - 1, cd["index"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["bulletin"]], sheet.Rows.Cells[row - 1, cd["bulletin"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["globalKb"]], sheet.Rows.Cells[row - 1, cd["globalKb"]]].Merge();
        }

        public void writePlatformOrComponentAndKB(string content)
        {
            string platformOrComponent, localKb = "";
            string pattern = @"\(\s?(KB)?\d+\s?\)";
            string tempContent = content;
            Match match;
            if ( (match = Regex.Match(content.Trim().ToUpper(), pattern)).Index  != 0) // we could not even found a parenthesis
            {
                platformOrComponent = content.Substring(0, match.Index);
                content = content.Substring(match.Index);
                localKb = content.Substring(1, content.IndexOf(")") - 1);
                sheet.Cells[row, myTable.getSystemOrComponent()] = utilities.removeExtraSpaces( utilities.deleteComments(platformOrComponent));   //table knows if the table has components or OS in the first row
                sheet.Cells[row, cd["kb"]] = localKb;                                
            }
            else
            {
                sheet.Cells[row, cd["kb"]] = get_Kb();
                sheet.Cells[row, myTable.getSystemOrComponent()] = tempContent.Trim();
            }
        }

        ~FoundstoneMethods() {
            utilities.releaseObject(application);
            utilities.releaseObject(workBook);
            utilities.releaseObject(workBooks);
            utilities.releaseObject(sheet);
            utilities.releaseObject(sheets);
        }

        public void set_Bulletin(string input)
        {
            bulletin = input;
        }

        public void set_Cvss(string input)
        {
            cvss = input;
        }

        public void set_Description(string input)
        {
            description = input;
        }

        public void set_Kb(string input)
        {
            kb = input;
        }

        public string getExcelResult()
        {
            return excelResult;
        }

        public bool isResultOpen()
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
    }
}
