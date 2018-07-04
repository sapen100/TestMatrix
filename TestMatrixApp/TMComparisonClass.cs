using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace TestMatrixApp
{
    class TMComparisonClass
    {
        public virtual void loadOldTMExcel(string OldTM) { }
        public virtual int numberBulletinsOldTM(){return -1;}
        public virtual void compareTMs(string oldTM, string newTM) { }
        public virtual Bulletin getBulletin(List<Bulletin> TM, Item globalKb) {return null;}
        public virtual void markEntireBulletinAsNew(Bulletin newBulletin){}
        public virtual void addDeletedBulletin(Bulletin deletedBulletin, int lastRow) {}
        public virtual void compareBulletin(Bulletin bulletinDoc, Bulletin bulletinWeb){}
        public virtual void verifyEquality(RowInTM rowItemDoc,RowInTM rowItemWeb){}
        public virtual RowInTM getRowItem(List<RowInTM> rowItems, string kb, string platform, string component){return null;}
        public virtual void changeColorCell(Item Value, Color color){}
        public virtual void changeColorRow(RowInTM rowItem, Color color) {}
        public virtual void changeColorCell(int row, int col, Color color){}
        public virtual List<Bulletin> loadExcel(string TM) {return null;}
        public virtual string getKBAt(int index) {return "";}
        public virtual void openExcel(string file){}
        public virtual void closeExcel(string file){}

        public virtual Bulletin getBulletin(List<Bulletin> TM, Item globalKb, Item cve) { return null; }    
    }
}
