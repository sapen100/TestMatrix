using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using excel = Microsoft.Office.Interop.Excel;

namespace TestMatrixApp
{    
    class TMComparisonSCAP : TMComparisonClass
    {        

        Dictionary<string, int> cd = new Dictionary<string, int>() { { "index", 1 }, { "bulletin", 2 }, { "cve", 3 }, { "description", 4 }, { "globalKb", 5 }, { "kb", 6 }, { "risk", 7 }, { "platform", 8 }, { "component", 9 }, { "status", 10 }, { "na", 13 }, { "supported", 14 }, { "bulletinsRep", 18 } };

        excel.Application application;
        excel.Workbooks workBooks;
        excel.Workbook workBook;
        excel.Sheets sheets;
        excel.Worksheet sheet;
        Color different, newItem, deleted;
        Utilities utilities;
        ErrorMessages em;
        List<Bulletin> TMDoc;

        public TMComparisonSCAP()
        {            
            em        = new ErrorMessages();
            utilities = new Utilities(em);
            different      = Color.Gold;
            newItem        = Color.Green;
            deleted        = Color.Red;
        }

        public override void loadOldTMExcel(string OldTM)
        {
            TMDoc = loadExcel(OldTM);
        }

        public override int numberBulletinsOldTM()
        {
            return TMDoc.Count;
        }


        //This method compares the two TM 
        //We know that the out is going to be either the OldTM or the NewTM, that is why, we are not gonna open it again
        public override void compareTMs(string oldTM, string newTM)
        {            
            List<Bulletin> TMWeb = loadExcel(newTM);
            openExcel(newTM);
            Bulletin bulletinDoc;
            int lastRow = 0;
            foreach(Bulletin bulletinWeb in TMWeb)
            {
                bulletinDoc = getBulletin(TMDoc, bulletinWeb.getGlobalKb());
                if (bulletinDoc != null)
                    compareBulletin(bulletinDoc, bulletinWeb);
                else {
                    markEntireBulletinAsNew(bulletinWeb);
                }
                lastRow = bulletinWeb.getBulletin().row + bulletinWeb.getRows().Count();
            }
            
            //Search for deleted bulletins
            Bulletin bulletinWeb_;
            int countDelBulletins = 0;
            foreach (Bulletin bulletinDoc_ in TMDoc)
            {
                bulletinWeb_ = getBulletin(TMWeb, bulletinDoc_.getGlobalKb());
                if (bulletinWeb_ == null)
                {
                    addDeletedBulletin(bulletinDoc_, lastRow + countDelBulletins);
                    countDelBulletins++;
                }
                
            }
            closeExcel(newTM);
        }



        public override Bulletin getBulletin(List<Bulletin> TM, Item globalKb) 
        {
            foreach (Bulletin bulletin in TM) 
            {
                if (bulletin.getGlobalKb().value.Equals(globalKb.value))
                    return bulletin;
            }
            return null;
        }

        public override void markEntireBulletinAsNew(Bulletin newBulletin)
        {
            Item CVE = newBulletin.getCVE();
            changeColorCell(CVE, newItem);
            

            Item bulletin = newBulletin.getBulletin();
            int col = bulletin.col;
            int row = bulletin.row;

            for (int i = 0; i < newBulletin.getRows().Count(); i++)
            {
                changeColorCell(row + i, col, newItem);
            }            
        }

        public override void addDeletedBulletin(Bulletin deletedBulletin, int lastRow)
        {
            Item bulletin = deletedBulletin.getBulletin();
            sheet.Cells[lastRow, cd["bulletin"]] = bulletin.value;
            changeColorCell(lastRow, cd["bulletin"], deleted);

            Item CVE = deletedBulletin.getCVE();            
            sheet.Cells[lastRow, cd["cve"]] = CVE.value;
            changeColorCell(lastRow, cd["cve"], deleted);

            Item description = deletedBulletin.getDescription();
            sheet.Cells[lastRow, cd["description"]] = description.value;
            changeColorCell(lastRow, cd["description"], deleted);

            Item globalKb = deletedBulletin.getGlobalKb();
            sheet.Cells[lastRow, cd["globalKb"]] = globalKb.value;
            changeColorCell(lastRow, cd["globalKb"], deleted);            
        }



        public override void compareBulletin(Bulletin bulletinDoc, Bulletin bulletinWeb)
        {         

            Item CVE = bulletinWeb.getCVE();
            if (! (CVE.value).Equals(bulletinDoc.getCVE().value)) {
                changeColorCell(CVE, different);
            }

            Item bulletin = bulletinWeb.getBulletin();
            int col       = bulletin.col;
            int row       = bulletin.row;
            /*if (bulletin.value.Equals("MS13-040"))
                col = col;*/
            if (!(bulletin.value).Equals(bulletinDoc.getBulletin().value))
            {
                for(int i = 0; i < bulletinWeb.getRows().Count(); i++)
                {
                    changeColorCell(row + i, col, different);
                }
            }

            Item description = bulletinWeb.getDescription();
            if (!(description.value).Equals(bulletinDoc.getDescription().value))
            {
                changeColorCell(description, different);
            }
            
            //This section look for differences that might exist
            List<RowInTM> rowItemsDoc = bulletinDoc.getRows();
            string kb, platform, component;
            RowInTM rowItemDocument;
            foreach (RowInTM rowItemWeb in bulletinWeb.getRows())
            {
                /*if (rowItemWeb.platform.value.Equals("Windows Server 2003 with SP2 for Itanium-based Systems"))
                    col = col;*/
                kb        = utilities.removeContentIn(rowItemWeb.KB.value, "[]");
                platform  = utilities.removeExtraSpaces(rowItemWeb.platform.value);
                component = utilities.removeExtraSpaces(rowItemWeb.component.value);                
                
                if ((rowItemDocument = getRowItem(rowItemsDoc, kb, platform, component)) != null)
                 {
                     verifyEquality(rowItemDocument, rowItemWeb);
                }
                else
                    changeColorRow(rowItemWeb, newItem);
            }

            //This seciont looks for deleted content
            string deletedContet = "";
            List<RowInTM> rowItemsWeb = bulletinWeb.getRows();
            foreach (RowInTM rowItemDoc in rowItemsDoc)
            {
                kb        = utilities.removeContentIn(rowItemDoc.KB.value, "[]");  // the kb contains some comments [][], we do not have to consider those to comapre 
                platform  = utilities.removeExtraSpaces(rowItemDoc.platform.value);
                component = utilities.removeExtraSpaces(rowItemDoc.component.value);
                if (getRowItem(rowItemsWeb, kb, platform, component) == null) //the key was not found, so it was deleted
                {
                    deletedContet += kb + "|" + platform + "|" + component + "\n";
                }                
            }
            if (!deletedContet.Equals(""))
            {
                sheet.Cells[CVE.row, cd["bulletinsRep"] + 1] = deletedContet;
                changeColorCell(CVE.row, cd["bulletinsRep"] + 1, deleted);
            }
        }

        public override void verifyEquality(RowInTM rowItemDoc, RowInTM rowItemWeb)
        {
            if (!rowItemDoc.status.value.Trim().ToUpper().Equals(rowItemWeb.status.value.Trim().ToUpper()))
                changeColorCell(rowItemWeb.status, different);

            if (!rowItemDoc.replacedBy.value.Trim().ToUpper().Equals(rowItemWeb.replacedBy.value.Trim().ToUpper()))
                changeColorCell(rowItemWeb.replacedBy, different);
        }

        public override RowInTM getRowItem(List<RowInTM> rowItems, string kb, string platform, string component)
        {
            bool flag1 = false, flag2 = false, flag3 = false;
            foreach(RowInTM rowItem in rowItems)
            {

                if ((utilities.removeContentIn(rowItem.KB.value, "[]").Trim().ToUpper().Equals(kb.ToUpper())) && (utilities.removeExtraSpaces(rowItem.platform.value).Trim().ToUpper().Equals(platform.ToUpper())) && (utilities.removeExtraSpaces(rowItem.component.value).Trim().ToUpper().Equals(component.ToUpper())))
                    return rowItem;             
                /*if (utilities.removeContentIn(rowItem.KB.value, "[]").Trim().ToUpper().Equals(kb.ToUpper())) 
                    flag1 = true;

                if(utilities.removeExtraSpaces(rowItem.platform.value).Trim().ToUpper().Equals(platform.ToUpper())) 
                    flag2 = true;
                if (utilities.removeExtraSpaces(rowItem.component.value).Trim().ToUpper().Equals(component.ToUpper()))
                    flag3 = true;

                if (flag1 && flag2 && flag3)
                    return rowItem;*/
            }
            return null;
        }


        public override void changeColorCell(Item Value, Color color)
        {
            /*if (Value.col == cd["status"])
                Value.row = Value.row;*/
            sheet.Cells[Value.row, Value.col].Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
        }


        public override void changeColorRow(RowInTM rowItem, Color color) 
        {
            foreach(Item item in rowItem) {
                changeColorCell(item, color);
            }
        }

        public override void changeColorCell(int row, int col, Color color)
        {
            sheet.Cells[row, col].Interior.Color = System.Drawing.ColorTranslator.ToOle(color);
        }

        public override List<Bulletin> loadExcel(string TM)
        {            
            openExcel(TM);
            List<Bulletin> TMResult = new List<Bulletin>();
            List<RowInTM> rowItems = new List<RowInTM>();
            int row            = 3;
            Item index       = new Item();
            Item bulletin    = new Item();
            Item CVE         = new Item();
            Item description = new Item();
            Item globalKb    = new Item();
            RowInTM rowItem;
            string value = sheet.Cells[row, cd["globalKb"]].Text;

            while (!(sheet.Cells[row, cd["globalKb"]].Text).Equals("")) {
                if (row == 3) //This is the first time in the loop
                {
                    index = new Item(sheet.Cells[row, cd["index"]].Text, row, cd["index"]);
                    bulletin    = new Item(sheet.Cells[row, cd["bulletin"]].Text, row, cd["bulletin"]);                    
                    CVE         = new Item(sheet.Cells[row, cd["cve"]].Text, row, cd["cve"]);
                    description = new Item(sheet.Cells[row, cd["description"]].Text, row, cd["description"]);
                    globalKb    = new Item(sheet.Cells[row, cd["globalKb"]].Text, row, cd["globalKb"]);
                }

                if (!globalKb.value.Equals(sheet.Cells[row, 5].Text)) // are we in a NEW BULLETIN?
                { 
                    TMResult.Add(new Bulletin(index, bulletin, CVE, description, globalKb, new List<RowInTM>(rowItems)));
                    rowItems    = new List<RowInTM>();
                    index       = new Item(sheet.Cells[row, cd["index"]].Text, row, cd["index"]);
                    bulletin    = new Item(sheet.Cells[row, cd["bulletin"]].Text, row, cd["bulletin"]);   
                    CVE         = new Item(sheet.Cells[row, cd["cve"]].Text, row, cd["cve"]);
                    description = new Item(sheet.Cells[row, cd["description"]].Text, row, cd["description"]);
                    globalKb    = new Item(sheet.Cells[row, cd["globalKb"]].Text, row, cd["globalKb"]);
                }
                
                rowItem            = new RowInTM();
                rowItem.KB         = new Item(sheet.Cells[row, cd["kb"]].Text, row, cd["kb"]);
                rowItem.risk       = new Item(sheet.Cells[row, cd["risk"]].Text, row, cd["risk"]);
                rowItem.platform   = new Item(sheet.Cells[row, cd["platform"]].Text, row, cd["platform"]);
                rowItem.component  = new Item(sheet.Cells[row, cd["component"]].Text, row, cd["component"]);
                rowItem.status     = new Item(sheet.Cells[row, cd["status"]].Text, row, cd["status"]);
                rowItem.replacedBy = new Item(sheet.Cells[row, cd["bulletinsRep"]].Text, row, cd["bulletinsRep"]);
                rowItems.Add(rowItem);

                row++;
            }
            
            TMResult.Add(new Bulletin(index, bulletin, CVE, description, globalKb, new List<RowInTM>(rowItems)));  //The last bulletin needs to be added
            closeExcel(TM);               
            return TMResult;
        }

        public override string getKBAt(int index) 
        {
            return TMDoc[index].getGlobalKb().value;
        }

        public override void openExcel(string file)
        {            
            if (File.Exists(file))
            {
                application = new excel.Application();
                application.DisplayAlerts = false;
                workBooks = application.Workbooks;
                workBook = workBooks.Open(file, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                sheets = workBook.Sheets;
                sheet = sheets.get_Item(1);                                
            }
        }


        public override void closeExcel(string file)
        {
            workBook.Save();
            workBook.Close(Microsoft.Office.Interop.Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing);
            application.Quit();    
        }        
    }
    

    class Bulletin
    {
        Item index;
        Item bulletin;
        Item CVE;
        Item description;
        Item globalKb;
        List<RowInTM> rows;

        public Bulletin(Item index, Item bulletin, Item CVE, Item description, Item globalKb, List<RowInTM> rows) 
        {
            this.index       = index;
            this.bulletin    = bulletin;
            this.CVE         = CVE;
            this.description = description;
            this.globalKb    = globalKb;
            this.rows        = rows;            
        }



        public Item getBulletin(){
            return bulletin;
        }

        public Item getCVE()
        {
            return CVE;
        }

        public Item getDescription()
        {
            return description;
        }

        public Item getGlobalKb()
        {
            return globalKb;
        }

        public List<RowInTM> getRows()
        {
            return rows;
        }

    }

    class RowInTM : IEnumerable<Item> {
        public Item KB;
        public Item risk;
        public Item platform;
        public Item component;
        public Item status;
        public Item replacedBy;


        public IEnumerator<Item> GetEnumerator()
        {
            yield return KB;
            yield return risk;
            yield return platform;
            yield return component;
            yield return status;
            yield return replacedBy;            
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public RowInTM() { 
        }


        public RowInTM(Item KB, Item risk, Item platform, Item component, Item status, Item replacedBy)
        {
            this.KB         = KB;
            this.risk       = risk;
            this.platform   = platform;
            this.component  = component;
            this.status     = status;
            this.replacedBy = replacedBy;            
        }
    }

    class Item{
        public string value;
        public int row;
        public int col;

        public Item(string value, int row, int col){
            this.value = value;
            this.row   = row;
            this.col   = col;
        }

        public Item(Item item) {
            this.value = item.value;
            this.row = item.row;            
            this.col = item.col;
        }

        public Item()
        {            
        } 
    }
}
