using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using excel = Microsoft.Office.Interop.Excel;

namespace TestMatrixApp
{
    class ScapMethods
    {
        //Global Variables        
        const string originalTemplate = "OriginalSCAPTM.xlsx";
        string       excelResult      = "TMResult.xlsx";
        string       configPath;
        int row, col, index, tableIniRow;

        Dictionary<string, int> cd = new Dictionary<string, int>() { { "index", 1 }, { "bulletin", 2 }, { "cve", 3 }, { "description", 4 }, { "globalKb", 5 }, { "kb", 6 }, { "risk", 7 }, { "platform", 8 }, { "component", 9 }, { "status", 10 }, { "na", 13 }, { "supported", 14 }, { "bulletinsRep", 18 } };

        public Table myTable;
        public Applications applications;
        public Utilities utilities;
        public ExcelUtilities excelUtils;        


        excel.Application application;
        excel.Workbooks workBooks;
        excel.Workbook workBook;
        excel.Sheets sheets;
        public excel.Worksheet sheet;


        string bulletin, cvss, description, kb;

        public ScapMethods(string applicationLocation, string configPath, ErrorMessages em, Utilities utilities)
        {
            this.configPath = configPath;
            excelResult = applicationLocation + "\\" + excelResult;
            this.utilities = utilities;
            myTable = new Table(cd["platform"], cd["component"], configPath, em);
            applications = new Applications(utilities, configPath, "SCAPApplications");
            excelUtils = new ExcelUtilities(applications, utilities, cd);
            row = 3;   //our template starts at this row            
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
        }
        

        //This method inserts operating systems to the TM if the application is described in Applications.xlsx
        public void insertSystems()
        {
            string platform = (sheet.Cells[row - 1, cd["platform"]].Text).Trim();
            if (platform.Equals("")) //This means that if it was an application the one that was inserted, we gotta see if there are platforms to insert for this application(e.g explorer 7.0)
            {
                string component               = (sheet.Cells[row - 1, cd["component"]].Text).Trim();                
                //ArrayList list = applications.getSystems(utilities.removeContentIn(component, "()"));
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
        public void insertSystem(string component, string system, string supported)
        {
            sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]].Text;
            sheet.Cells[row, cd["risk"]] = sheet.Cells[row - 1, cd["risk"]].Text;
            sheet.Cells[row, cd["bulletinsRep"]] = sheet.Cells[row - 1, cd["bulletinsRep"]].Text;
            sheet.Cells[row, cd["platform"]] = system;
            sheet.Cells[row, cd["component"]] = component;
            sheet.Cells[row, cd["supported"]] = supported;
            if (supported.Trim().ToUpper().Equals("NO"))
                sheet.Cells[row, cd["status"]] = "Not applicable";
            row++;

        }

        //This method is a insertSystems method helper that inserts a system if needed.
        public void insertSystemSameLine(string component, string system, string supported)
        {
            sheet.Cells[row - 1, cd["platform"]] = system;
            sheet.Cells[row - 1, cd["supported"]] = supported;
            if (supported.Trim().ToUpper().Equals("NO"))
                sheet.Cells[row - 1, cd["status"]] = "Not applicable";
        }

        private void writeSuported()
        {
            if (!(myTable.isHeadSupported()))
            {
                setNotApplicable();
                sheet.Cells[row, cd["platform"]] = sheet.Cells[row, cd["platform"]].Text + "(CORE)";
                return;
            }

            if (myTable.isSystemNotSupported(sheet.Cells[row, cd["platform"]].Text.ToUpper()))
            {
                setNotApplicable();
                return;
            }

            if (myTable.isApplicationNotSupported(sheet.Cells[row, cd["component"]].Text.ToUpper()))
            {
                setNotApplicable();
                return;
            }
            else
                sheet.Cells[row, cd["supported"]] = "Yes";
        }

        private void setNotApplicable()
        {
            sheet.Cells[row, cd["supported"]] = "No";
            sheet.Cells[row, cd["status"]] = "Not applicable";
        }





        //This method is trigered whenever the table we are working has 4 columns and it is a OPERATYING SYSTEM TABLE , that means that the only column that many contain mone than one content is the PLATFORM column
        public void writeSystem(string content, string bulletinsReplacedBy)
        {
            string comments = "";
            string contentComment;
            string shellContent;
            //string[] systems = content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            string[] systems = utilities.splitContent(content);

            if (systems.Length == 1)
            {
                if (!(systems[0].Trim().ToUpper().Equals("NOT APPLICABLE")))
                {
                    comments = utilities.extractComments(systems[0]);
                    sheet.Cells[row, cd["platform"]] = utilities.deleteComments(systems[0]);
                    shellContent = sheet.Cells[row, cd["kb"]].Text;
                    sheet.Cells[row, cd["kb"]] = shellContent;
                    sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsReplacedBy, sheet.Cells[row, cd["kb"]].Text);
                }
                writeSuported();
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
                        sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsReplacedBy, sheet.Cells[row, cd["kb"]].Text);
                        writeSuported();
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


        //This method is trigered whenever the table we are working has 5 columns and it is a SYSTEM TABLE , that means that the only column that many contain mone than one content is the COMPONENT column
        public void writeComponent(string content, string bulletinsReplacedBy)
        {
            string comments = "";
            string[] components = utilities.splitContent(content);

            if (components.Length == 1)
            {
                if (!(components[0].Trim().ToUpper().Equals("NOT APPLICABLE")))
                {
                    comments = utilities.extractComments(components[0]);
                    sheet.Cells[row, cd["component"]] = utilities.deleteComments(components[0]);
                    sheet.Cells[row, cd["kb"]] = sheet.Cells[row, cd["kb"]].Text + comments;
                    sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsReplacedBy, sheet.Cells[row, cd["kb"]].Text);
                }
                writeSuported();
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
                        sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsReplacedBy, sheet.Cells[row, cd["kb"]].Text);
                        writeSuported();
                        row++;
                        separateSystemsOrComponentsInTwoIfRequired();
                    }
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show("Error Writing Components: " + e.Message + "\n\n Please verify if the following data is consistent: " + getComponents(components));
                }
            }
        }

        //This method is trigered whenever the content has more than one item in it, so it needs to be splitted.
        // @_otherCellContent is used to extract comments, that may be in the other column, and also to fix some Microsoft errors. (the content is not necesarily the one where the KB is located. it could also be in the other cell, that is what @othercellcontent contains)
        public void writeComponentsOnly(string content, string bulletinsReplacedBy, string otherCellContent = "")
        {
            string comments = "";
            string cleanContent = "";
            string otherCellComments = utilities.extractComments(otherCellContent);
            otherCellContent = utilities.deleteComments(otherCellContent);

            //string[] components = utilities.deleteCommentsInTheMiddle(content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries));
            //string[] components = utilities.deleteCommentsInTheMiddle(utilities.splitBasedOnContent(content));
            string[] components = utilities.splitContent(content);

            if (components.Length == 1) //is there more than one component in the cell
            {
                cleanContent = utilities.deleteComments(components[0]);
                sheet.Cells[row, cd["component"]] = cleanContent;
                sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsReplacedBy, sheet.Cells[row, cd["kb"]].Text);
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
                    sheet.Cells[row, cd["risk"]] = risk;
                    sheet.Cells[row, cd["kb"]] = utilities.deleteParenthesis(components[i + 1]) + comments;
                    sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsReplacedBy, sheet.Cells[row, cd["kb"]].Text);
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
            if (utilities.moreThanOneContent(sheet.Cells[row - 1, cd["platform"]].Text))    //if the column platform has more than once platform, we separate them here
            {
                string strSystems = sheet.Cells[row - 1, cd["platform"]].Text;
                string[] systems = utilities.separateAndOccurrences(strSystems);
                sheet.Cells[row - 1, cd["platform"]] = systems[0];
                sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]];
                sheet.Cells[row, cd["risk"]] = sheet.Cells[row - 1, cd["risk"]];
                sheet.Cells[row, cd["component"]] = sheet.Cells[row - 1, cd["component"]];
                sheet.Cells[row, cd["supported"]] = sheet.Cells[row - 1, cd["supported"]];
                sheet.Cells[row, cd["platform"]] = systems[1];
                row++;
            }

            else if (utilities.moreThanOneContent(sheet.Cells[row - 1, cd["component"]].Text))  //if the column component has more than once component, we separate them here
            {
                string strComponents = sheet.Cells[row - 1, cd["component"]].Text;
                string[] components = utilities.separateAndOccurrences(strComponents);
                sheet.Cells[row - 1, cd["component"]] = components[0].Trim() + utilities.getBitEdition(sheet.Cells[row - 1, cd["component"]].Text);

                string tmpPlatform = sheet.Cells[row - 1, cd["platform"]].Text; //Before it starts putting more platforms, we copy what is the original

                insertSystems();

                sheet.Cells[row, cd["kb"]] = sheet.Cells[row - 1, cd["kb"]];
                sheet.Cells[row, cd["risk"]] = sheet.Cells[row - 1, cd["risk"]];
                sheet.Cells[row, cd["supported"]] = sheet.Cells[row - 1, cd["supported"]];
                sheet.Cells[row, cd["platform"]] = tmpPlatform;
                sheet.Cells[row, cd["component"]] = components[1].Trim();
                row++;
            }
            insertSystems();
        }
        



        //Once we completed traversing a word Document, we write the text that we got from the word and merge the cells that need to be merged
        public void documentEnd()
        {
            if (tableIniRow != row) //nothing was add, so there is nothing to merge
            {
                sheet.Cells[tableIniRow, cd["index"]] = get_Index().ToString();
                //sheet.Cells[tableIniRow, cd["bulletin"]] = get_Bulletin();
                sheet.Cells[tableIniRow, cd["cve"]] = get_Cvss();
                sheet.Cells[tableIniRow, cd["description"]] = get_Description();
                //sheet.Cells[tableIniRow, cd["globalKb"]] = get_Kb();
                mergeCells();
                tableIniRow = row;
                myTable.setSupportedHeader(true);   //We assume that eveything is supported until we ran into a header that is not supported. 
            }
        }            

        //Save the excel created
        public void saveDocument()
        {
            workBook.Save();
            workBook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            application.Quit();
        }

        //Merge cells
        private void mergeCells()
        {
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["index"]], sheet.Rows.Cells[row - 1, cd["index"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["cve"]], sheet.Rows.Cells[row - 1, cd["cve"]]].Merge();
            sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["description"]], sheet.Rows.Cells[row - 1, cd["description"]]].Merge();
            string bulletin = get_Bulletin();
            string globalKb = get_Kb();
            for (int i = tableIniRow; i < row; i++)
            {
                sheet.Cells[i, cd["bulletin"]] = bulletin;
                sheet.Cells[i, cd["globalKb"]] = globalKb;
                //sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["bulletin"]], sheet.Rows.Cells[row - 1, cd["bulletin"]]].Merge();
                //sheet.Rows.Range[sheet.Rows.Cells[tableIniRow, cd["globalKb"]], sheet.Rows.Cells[row - 1, cd["globalKb"]]].Merge();            
            }
        }

        //This method writes the KB and the platform or component.
        public void writePlatformOrComponentAndKB(string content, string bulletinsRep)
        {
            int endPos, initPos;
            string comments = utilities.extractComments(content);
            string localPlatform = "";
            string localKb = "";
            string tempContent = content;

            if ((initPos = utilities.obtainInitKB(content)) != -1)
            {

            //int iniPos = content.IndexOf("(KB");
            //if (iniPos != -1)
            //{
                localPlatform = content.Substring(0, initPos);
                content = content.Substring(initPos);
                endPos = content.IndexOf(")");
                if (endPos != -1)
                {
                    localKb = content.Substring(1, endPos - 1);
                    sheet.Cells[row, myTable.getSystemOrComponent()] = utilities.deleteComments(localPlatform);   //table knows if the table has components or OS in the first row
                    sheet.Cells[row, cd["kb"]] = localKb + comments;
                    sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsRep, sheet.Cells[row, cd["kb"]].Text);
                    return;
                }
            }
            sheet.Cells[row, cd["kb"]] = get_Kb() + comments;
            sheet.Cells[row, cd["bulletinsRep"]] = utilities.getBulletinReplacedBy(bulletinsRep, sheet.Cells[row, cd["kb"]].Text);
            sheet.Cells[row, myTable.getSystemOrComponent()] = utilities.deleteComments(tempContent);
        }


          //Destructor
        ~ScapMethods()
        {
            utilities.releaseObject(application);
            utilities.releaseObject(workBook);
            utilities.releaseObject(workBooks);
            utilities.releaseObject(sheet);
            utilities.releaseObject(sheets);
        }        

        public  void set_Bulletin(string input)
        {
            bulletin = input;
        }

        public void set_Cvss(string input)
        {
            cvss = input;
        }
        public  void set_Description(string input)
        {
            description = input;
        }
        public void set_Kb(string input)
        {
            kb = input;
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

        public string getExcelResult()
        {
            return excelResult;
        }

        public int getRow() {
            return row;
        }

        public int getCdAt(string value)
        {
            return cd[value];
        }

        public string getFromSheet(int r, int c)
        {
            return (sheet.Cells[r, c].Text).Trim();
        }
    }
}
