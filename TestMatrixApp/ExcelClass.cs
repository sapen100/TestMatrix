using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HtmlAgilityPack;

namespace TestMatrixApp
{
    class ExcelClass
    {
        public virtual void newDocument() { }       
        public virtual void set_Bulletin(string input) { }
        public virtual void set_Description(string input) { }
        public virtual void set_Kb(string input) { }
        public virtual void documentEnd() { }        
        public virtual string getExcelResult(){
            return ""; 
        }
        public virtual bool isResultOpen() { return true; }// Checks whether Result.xlsx is open
        public virtual void saveDocument() { }
        public virtual void readTablesFromWord(string wordFile) { }
        public virtual void readTablesFromWeb(HtmlDocument wordFile) { }
    }
}
