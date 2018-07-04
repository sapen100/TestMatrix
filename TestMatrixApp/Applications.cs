using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using Novacode;
using excel = Microsoft.Office.Interop.Excel;

namespace TestMatrixApp
{
    //This class was designed to be in charge of storing the applications with their respective operating systems that are supported in    
    class Applications
    {
        string sheetName;
        string excelDocument;
        Dictionary<string, ArrayList> components;
        excel.Application excelApplication;
        excel.Workbooks workBooks;
        excel.Workbook workBook;
        excel.Sheets sheets;
        excel.Worksheet sheet;
        Utilities utilities;
        

        public Applications(Utilities utilities, string excelDocument, string sheetName) {
            this.utilities     = utilities;
            this.excelDocument = excelDocument;
            this.sheetName     = sheetName;
            loadDataFDocument();
        }

        private void loadDataFDocument(){
            if (File.Exists(excelDocument))
            {
                excelApplication = new excel.Application();
                excelApplication.DisplayAlerts = false;
                workBooks = excelApplication.Workbooks;
                workBook = workBooks.Open(excelDocument, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheets = workBook.Sheets;
                sheet = sheets.get_Item(sheetName);
                loadDataToDictionary();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("ERROR: Excel Config File for TM was not found file: " + excelDocument);
            }
        }
        
            

        //Copies information from the worksheet of the Applications to memory
        private void loadDataToDictionary(){
            components = new Dictionary<string, ArrayList>();
            int row = 2;
            string component, os, supported;
            ArrayList oss;
            //System.Windows.Forms.MessageBox.Show(sheet.Cells[1, 1].Text);
            while ( !((sheet.Cells[row, 1].Text).Equals("")) ){
                component = utilities.removeExtraSpaces(sheet.Cells[row, 1].Text.Trim());
                os        = utilities.removeExtraSpaces(sheet.Cells[row, 2].Text.Trim());
                supported = utilities.removeExtraSpaces(sheet.Cells[row, 3].Text.Trim());
                if (components.ContainsKey(component))
                {
                    oss = components[component];
                    oss.Add(os);
                    oss.Add(supported);
                    components.Remove(component);
                    components.Add(component, oss);                    
                }
                else {
                    oss = new ArrayList();
                    oss.Add(os);
                    oss.Add(supported);
                    components.Add(component, oss);
                }
                row++;
            }            
            excelApplication.Quit();
        }

        public ArrayList getSystems(string input)
        {            
            if (components.ContainsKey(input))
                return components[input];
            return null;
        }

        private void showtest() {
            foreach (string x in components.Keys) {
                System.Windows.Forms.MessageBox.Show(x );
                ArrayList a = components[x];
                for (int i = 0; i < a.Count; i=i+2 )
                {
                    System.Windows.Forms.MessageBox.Show(x+ "|"+ a[i] + " yesno=" + a[i+1] );
                }
            }
        }

        public bool applicationExist(string input) 
        {
            return components.ContainsKey(input);
        }

        ~Applications() {
            /*System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBooks);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheets);
           excelApplication.Quit();*/            
        }
    }
}
