using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Novacode;
using System.IO;
using HtmlAgilityPack;
using excel = Microsoft.Office.Interop.Excel;

namespace TestMatrixApp
{

    //This class was designed to map all the attributes of a table that appears in the bulletin.
    
    class Table
    {
        int sysOrComponent;
        int colPlatform;
        int colComponent;
        string strSystemOrComponent;
        string headersFile;
        string headerSheet        = "ValidHeaders";
        string headerNotSupported = "SCAPNotSupported";
        List <List<string>> headers;
        ErrorMessages em;
        

        //supported stuff
        List<string> supportedHeaders;
        List<string> supportedSystems;
        List<string> supportedApps;
        bool headerSupported;

        //Variables to manage Excel
        excel.Application application;
        excel.Workbooks   workBooks;
        excel.Workbook    workBook;
        excel.Sheets      sheets;
        excel.Worksheet   sheet;
        

        public Table(int platform, int component, string headersConfiguration, ErrorMessages em)
        {
            this.em         = em;
            colPlatform     = platform;
            colComponent    = component;
            headerSupported = true;
            headersFile     = headersConfiguration;
            openExcelFile();
                loadHeaders();
                loadSupported();
            closeExcelFile();
        }

        private void openExcelFile()
        {
            if (File.Exists(headersFile))
            {
                application = new excel.Application();
                application.DisplayAlerts = false;
                workBooks = application.Workbooks;
                workBook = workBooks.Open(headersFile, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheets = workBook.Sheets;
               
            }
            else {
                em.showErrorFileNotExist(headersFile);
                System.Environment.Exit(0);                
            }
        }

        //This method reads from a sheet all the valid table headers that are gonna be taken into account 
        private void loadHeaders()
        {
            try
            {
                sheet = sheets.get_Item(headerSheet);   //Sheet 3 contains the valid headers
            }
            catch {
                em.showErrorSheetNotExist(headersFile, headerSheet);                                 
            }
            headers = new List<List<string>>();
            int row = 1;
            int col;
            List<string> rowList;
            while (!((sheet.Cells[row, 1].Text).Equals("")))
            {
                col = 1;
                rowList = new List<string>();
                while (!((sheet.Cells[row, col].Text).Equals("")))
                {
                    rowList.Add(sheet.Cells[row, col].Text());
                    col++;
                }
                headers.Add(rowList);
                row++;
            }                                      
        }

        //This method reads from a sheet all the supported headers, systems and applications  
        private void loadSupported()
        {
            try
            {
                sheet = sheets.get_Item(headerNotSupported);   //Sheet 2 contains the valid headers
            }
            catch
            {
                em.showErrorSheetNotExist(headersFile, headerNotSupported);
            }
            
            int headerCol = 1;
            int systemCol = 2;
            int appCol    = 3;            
            supportedHeaders = new List<string>();
            supportedSystems = new List<string>();
            supportedApps    = new List<string>();
            string value;
            int row = 2;            
            while ( (!((sheet.Cells[row, headerCol].Text).Equals(""))) || (!((sheet.Cells[row, systemCol].Text).Equals(""))) || (!((sheet.Cells[row, appCol].Text).Equals(""))) )
            {
                value = (sheet.Cells[row, headerCol].Text).Trim();
                if (!((sheet.Cells[row, headerCol].Text).Trim().Equals("")))
                    supportedHeaders.Add(value.Trim().ToUpper());

                value = (sheet.Cells[row, systemCol].Text).Trim();
                if (!((sheet.Cells[row, systemCol].Text).Trim().Equals("")))
                    supportedSystems.Add(value.Trim().ToUpper());

                value = (sheet.Cells[row, appCol].Text).Trim();
                if (!((sheet.Cells[row, appCol].Text).Trim().Equals("")))
                    supportedApps.Add(value.Trim().ToUpper());
                row++;
            }
           
        }

        private void closeExcelFile()
        {
            workBook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            application.Quit();
        }

        //Testing function 
        private void showList() { 
            foreach(List<string> l in headers){
                foreach(string ele in l){
                    System.Windows.Forms.MessageBox.Show(ele);
                }
            }
        }

        //Given a row with columns, verify whether that row defines a table with affected software. 
        public bool containsAffectedSoftware(Novacode.Row firstRow)
        {

            if ((firstRow.Cells.Count == 4) || (firstRow.Cells.Count == 5))
            {
                if (isValidHeader(firstRow))
                {
                    return true;
                }

            }
            return false;
        }        

        private bool equalToHeader(Novacode.Row row, string[] validHeader)
        {
            for (int i = 1 ; i < 4 ; i++)
            {
                //                System.Windows.Forms.MessageBox.Show("comparando:" + validHeader[i] + ";" + row.Cells[i].Paragraphs[0].Text);
                if (!(validHeader[validHeader.Length-i].Equals(((row.Cells[row.Cells.Count-i].Paragraphs[0].Text).Trim()).ToUpper())))
                {
                    return false;
                }
            }
            return true;
        }

        public string getCVS(Novacode.Table table)
        {
            string tableHeaderContainCVE = "VULNERABILITY SEVERITY RATING AND MAXIMUM SECURITY IMPACT BY AFFECTED SOFTWARE";
            if ((table.Rows[0].Cells.Count == 1) && (tableHeaderContainCVE.Equals(table.Rows[0].Cells[0].Paragraphs[0].Text.ToUpper().Trim())))
            {
                return concatCVEs(table.Rows[1]);
            }
            return "";
        }

        public string getCVS(HtmlNodeCollection row)
        {
            string result = "";
            foreach(HtmlNode col in row)
            {                                
                if (col.InnerText.Trim().ToUpper().Contains("CVE-"))
                {
                    result += getCVEfromString(col.InnerHtml) + " ";
                }
            }
            return result;
        }

        public bool tableContainsCVEs(HtmlNodeCollection row)
        {

            string tableHeaderContainCVE = "VULNERABILITY SEVERITY RATING AND MAXIMUM SECURITY IMPACT BY AFFECTED SOFTWARE";            
            if ( (row.Count == 1) && (tableHeaderContainCVE.Equals(row[0].InnerHtml.ToUpper().Trim())) )
            {                
                return true;
            }            
            return false;
        }

        public bool isCVEHeader(string value)
        {

            string tableHeaderContainCVE = "VULNERABILITY SEVERITY RATING AND MAXIMUM SECURITY IMPACT BY AFFECTED SOFTWARE";
            if (tableHeaderContainCVE.Equals(value.ToUpper().Trim()))
            {
                return true;
            }
            return false;
        }


        public bool tableContainsCVEs(Novacode.Row row)
        {

            string tableHeaderContainCVE = "VULNERABILITY SEVERITY RATING AND MAXIMUM SECURITY IMPACT BY AFFECTED SOFTWARE";
            if ((row.Cells.Count == 1) && (tableHeaderContainCVE.Equals(row.Cells[0].Paragraphs[0].Text.ToUpper().Trim())))
            {
                return true;
            }
            return false;
        }  

        private string concatCVEs(Novacode.Row row)
        {
            string result = "";
            for (int i = 0; i < row.Cells.Count; i++)
            {
                //TODO maybe there are more than on paragraph
                Paragraph paragrap = row.Cells[i].Paragraphs[0];
                if (paragrap.Text.Contains("CVE-"))
                {
                    result += getCVEfromString(paragrap.Text) + " ";
                }
            }
            return result;
        }

        private string getCVEfromString(string line)
        {
            int inipos = line.IndexOf("CVE-");
            if (inipos > 0)
            {
                return line.Substring(inipos);
            }
            return "";
        }


        private bool isValidHeader(Novacode.Row firstRow)
        {                                      
            bool flag;
            foreach(List<string> list in headers){
                if (list.Count() == firstRow.Cells.Count)
                {
                    flag = true;
                    for (int i = 0; i< list.Count; i++)
                    {
                        if (!(list.ElementAt(i).Equals((firstRow.Cells[i].Paragraphs[0].Text).Trim().ToUpper())))
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                        return true;
                }
            }
            return false;                                     
        }

        public bool isValidHeader(HtmlNodeCollection header)
        {                                      
            bool flag;
            foreach(List<string> list in headers){
                if (list.Count() == header.Count)
                {
                    flag = true;
                    for (int i = 0; i< list.Count; i++)
                    {
                        if (!(list.ElementAt(i).Equals((header[i].InnerText).Trim().ToUpper())))
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                        return true;
                }
            }
            return false;                                     
        }
        


        public void setSystemOrComponent(string value)
        {
            /*if (rowContent.Cells.Count == 4)
            {*/                
                //if (rowContent.Cells[0].Paragraphs[0].Text.ToUpper().Trim().Equals("OPERATING SYSTEM"))
            if (value.ToUpper().Trim().Equals("OPERATING SYSTEM"))
            {
                setSystemOrComponent(colPlatform);
                strSystemOrComponent = "os";
            }
            else
            {                
                setSystemOrComponent(colComponent);
                strSystemOrComponent = "component";
            }


            /*switch (rowContent.Cells[0].Paragraphs[0].Text)
            {                  
                case "Operating System": setSystemOrComponent(colPlatform);
                                         strSystemOrComponent = "os";
                    break;
                case "Software": setSystemOrComponent(colComponent);
                    break;
                case "Developer Tools":                                  setSystemOrComponent(colComponent);
                    break;
                case "Microsoft Office and Other Productivity Software": setSystemOrComponent(colComponent);
                    break;                        
                default: throw new Exception("ERROR: CLASS TABLE: Trying to parse the header, but the table header is not cataloged");
            }
        }

        if (rowContent.Cells.Count == 5)
        {
            setSystemOrComponent(colComponent);
            strSystemOrComponent = "component";
            switch (rowContent.Cells[0].Paragraphs[0].Text)
            {
                case "Operating System": setSystemOrComponent(colPlatform);
                    strSystemOrComponent = "os";
                    break;                    
                //case "Software": setSystemOrComponent(colComponent);
                  //               strSystemOrComponent = "component";
                    //break;                    
                //default: strSystemOrComponent = "component";//throw new Exception("ERROR: CLASS TABLE: Trying to parse the header, but the table header is not cataloged");
            }                
        }*/

        }

        //Returns the col where we should put the data, if it is a system, we have to return column to write the platform in 
        public int getSystemOrComponent() {
            return sysOrComponent;
        }
        
        public bool isSystem()
        {                
            return (strSystemOrComponent == "os");
        }
        
        public bool isComponent()
        {
            return (strSystemOrComponent == "component");
        }

        private void setSystemOrComponent(int value)
        {
            sysOrComponent = value;
        }

        public void setHeaderSupporteTrue()
        {
            headerSupported = true;
        }

        public void setSupportedHeader(string input) 
        {
            if (!(input.Equals("")))
            {
                input = input.Trim().ToUpper();
                foreach (string value in supportedHeaders)
                {
                    if (value.Equals(input)) {
                        headerSupported =  false;
                        return;
                    }
                }                
            }
            headerSupported = true; ;
        }

        public void setSupportedHeader(bool value){
            headerSupported = value;
        }

        public bool isSystemNotSupported(string input) 
        {
            if (!(input.Equals("")))
            {
                input = input.Trim().ToUpper();
                foreach (string value in supportedSystems)
                {
                    if (input.Contains(value)) {
                        return true;
                    }
                }                
            }
            return false;
        }

        public bool isApplicationNotSupported(string input)
        {
            if (!(input.Equals("")))
            {
                input = input.Trim().ToUpper();
                foreach (string value in supportedApps)
                {
                    if (input.Contains(value))
                    {
                        return true;
                    }
                }
            }
            return false;
        }


        public bool isHeadSupported() {
            return headerSupported;
        }
    }

}