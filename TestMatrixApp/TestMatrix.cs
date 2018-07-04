using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Text.RegularExpressions;
//using Microsoft.Office.Interop.Word;
using Novacode;
using HtmlAgilityPack;


namespace TestMatrixApp
{
    class TestMatrix
    {      
        string currentDir;
        string configPath;
        string sharePointURL    = "http://2010.corp.mcafee.com/sites/avertqa/rc_coordination/Processes%20Library/Other/TestMatrixConfig/TMConfig.xlsx";
        string configFileName   = "TMConfig.xlsx";
        string configFileNameSP = "TMConfigSP.xlsx";

        Dictionary<string, HtmlDocument> webPages;
        ExcelClass        excel;        
        Utilities         utilities;
        WebUtilities      webUtilities;    
        ScapMethods       scap;
        FoundstoneMethods foundstone;
        ErrorMessages     em;
        TMComparisonClass tmc;
        //TMComparisonSCAP  tmcScap;        
        
        public TestMatrix()
        {            
            currentDir      = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            em              = new ErrorMessages();
            this.utilities  = new Utilities(em);
            webUtilities    = new WebUtilities();
            webPages        = new Dictionary<string, HtmlDocument>();                     
        }
        

        public void processWordDocument(string wordFile)
        {            
            newDocument();
            if (File.Exists(wordFile))
            {                         
                readTextFromWord(wordFile);
                readTablesFromWord(wordFile);                        
            }            
            documentEnd();            
        }


        public bool downloadWebDocument(string url)
        {
            HtmlDocument page = new HtmlDocument();
            string title, kb;
            if (utilities.existWeb(url, ref page))
            {
                title = webUtilities.getTagContent(page, "title");                
                kb = utilities.getKB(title);
                webPages.Add(kb, page);
                return true;
            }
            return false;
        }
        
        public void processWebDocument(string url)
        {
            HtmlDocument page = new HtmlDocument();
            newDocument();
            if (utilities.existWeb(url, ref page))
            {
                readTextFromWeb(page);
                readTablesFromWeb(page);       
            }            
            documentEnd();
        }

        private void processWebDocument(HtmlDocument page)
        {            
            newDocument();         
            readTextFromWeb(page);
            readTablesFromWeb(page);            
            documentEnd();
        }

        public void processWebDocuments(string oldTM)
        {
            string kbDoc;
            tmc.loadOldTMExcel(oldTM);
            List<string> kbsInDocTM = new List<string>();
            int i = 0;
            int index = 0;
            while (i < tmc.numberBulletinsOldTM()) {
                kbDoc = utilities.getOnlyNumericValue(tmc.getKBAt(index));                
                if (webPages.ContainsKey(kbDoc))
                {
                    processWebDocument(webPages[kbDoc]);
                    kbsInDocTM.Add(kbDoc);
                }                
                i++;
                index++;
            }

            foreach (string key in webPages.Keys)
            {
                if (!kbsInDocTM.Contains(key))
                    processWebDocument(webPages[key]);
            }
        }

        //Ends the files that are being used. It returns true, when it was able to close applications, returns false if nothing was closed
        public bool endDocuments() {
            if (excel.isResultOpen())
            {
                saveDocument();
                deleteSPFileCopied();                
                return true;
            }
            else
                return false;
        }

        public void openResult()
        {
            System.Diagnostics.Process.Start(getExcelResult());
        } 
        


        private void deleteSPFileCopied() {
            if (File.Exists(currentDir + "\\" + configFileNameSP))
            {                
                File.Delete(currentDir + "\\" + configFileNameSP);
            }
        }

        private void readTextFromWord(string wordFilein)
        {
            try
            {
                string paragraph, bulletin, description, kb;
                using (DocX document = DocX.Load(wordFilein))
                {
                    int counter = 0;
                    while ((paragraph = document.Paragraphs[counter].Text).Equals("")) //If there are spaces at the beginning of the document
                        counter++;
                    bulletin = utilities.getBulletin(paragraph);
                    paragraph = document.Paragraphs[++counter].Text;
                    description = utilities.getDescription(paragraph);
                    kb = utilities.getKB(paragraph);
                    writeWordInformation(bulletin, description, kb);
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error Reading Word Document: " + e.Message + "\n Please close the document " + wordFilein + " and restart the execution");
                excel.saveDocument();
                System.Environment.Exit(0);
            }
        }

        private void readTextFromWeb(HtmlDocument page)
        {
            try
            {
                string title, bulletin, description, kb;                                
                title       = webUtilities.getTagContent(page, "title");
                bulletin    = utilities.getBulletin(title);                
                description = utilities.getWebDescription(title);
                kb          = utilities.getKB(title);
                writeWordInformation(bulletin, description, kb);                
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Error Reading Word Document: " + e.Message + "\n Please close the document " + page + " and restart the execution");
                excel.saveDocument();
                System.Environment.Exit(0);
            }
        }        

        public void setTMType(string input, bool local) 
        {
            switch (input)
            {
                case "FOUNDSTONE":
                    foundstone = new FoundstoneMethods(currentDir, configPath, em, utilities);
                    tmc        = new TMComparisonFoundstone();
                    if (local)
                        excel = new ExcelFoundstone(foundstone);
                    else
                        excel = new WebFoundstone(foundstone);
                    break;
                case "SCAP":
                        scap = new ScapMethods(currentDir, configPath, em, utilities);
                        tmc  = new TMComparisonSCAP();
                        if (local)
                            excel = new ExcelScap(scap);
                        else
                            excel = new WebScap(scap);
                    break;
                case "RM":  excel = new ExcelRM(currentDir, configPath, em, utilities);
                    break;                    
            }
        }        

        public void compareTMs(string oldTM, string newTM)
        {
            tmc.compareTMs(oldTM, newTM);
        }

        private void newDocument()
        {
            excel.newDocument();
        }


        private void writeWordInformation(string bulletin, string description, string kb)
        {
            excel.set_Bulletin(bulletin);
            excel.set_Description(description);
            excel.set_Kb("KB" + kb);                        
        }

        private void documentEnd() {        
            excel.documentEnd();            
        }
        

        //It defines whether the application will run locally or remotely
        public bool defineconfigPath(bool useSP){            
            if (useSP)
            {
                configPath = currentDir + "\\" + configFileNameSP;                
                return utilities.copyFileFromSharePoint(sharePointURL, configPath);
                
            }
            else {
                configPath = currentDir + "\\" + configFileName;    // We expect to find the TMConfig locally                                                
                if (!File.Exists(configPath))
                {
                    em.showErrorFileNotExist(configPath);                    
                    return false;
                }
                return true;
            }                
        }

        private string getExcelResult() {            
            return excel.getExcelResult();            
        }
        

        private void saveDocument() {            
            excel.saveDocument();            
        }        

        private void readTablesFromWord(string wordFile) {            
            excel.readTablesFromWord(wordFile);            
        }

        private void readTablesFromWeb(HtmlDocument page)
        {
            excel.readTablesFromWeb(page);
        }

        public string getSharepointUrl() 
        {
            return sharePointURL;
        }    

        ~TestMatrix()
        {
           // excel = null;
        }


    }
}
