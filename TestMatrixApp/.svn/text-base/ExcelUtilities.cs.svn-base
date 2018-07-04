using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using excel = Microsoft.Office.Interop.Excel;

namespace TestMatrixApp
{
    class ExcelUtilities
    {
        Applications applications;
        excel.Worksheet sheet;
        Dictionary<string, int> cd;
        
        Utilities utilities;

        public ExcelUtilities(Applications _applications, Utilities _utilities, Dictionary<string, int> _cd )        
        {
            applications = _applications;
            utilities = _utilities;
            cd = _cd;
        }

        //@_cd brings the structure of the TMTemplate column
        public void setSheet(excel.Worksheet _sheet)
        {
            sheet = _sheet;            
        }

        //-------------------------- THIS METHODS ARE COMMON IN FOUNDSTONE, SCAP AND RM-------------------------------------------------------------

        //This is the method that fixes the exception in which the real content is not where the KB is located
        public void replaceComponentContentIfRequired(string originalContent, string otherContent, int currentRow)
        {            
            if (!otherContent.Equals("")) //if this is the case, it means that it came from sent by a content that was on the the first column
            {
                int row = currentRow;
                string bitEdition   = utilities.getBitEdition(originalContent);
                string[] components = utilities.separateAndOccurrences(originalContent);

                for (int i = 0; i < components.Length; i++)
                {
                    if (i + 1 == components.Length)
                        bitEdition = "";

                    if (verifyIfOtherContentExistInApps(otherContent, components[i])) 
                    {
                        sheet.Cells[row - 1, cd["component"]] = otherContent;
                        return;
                    }                        
                }
            }
        }

        //This function verifies that the content of the other column is the valid one
        private bool verifyIfOtherContentExistInApps(string otherColumnContent, string originalContent)
        {
            string bitEdition = "";
            if (utilities.moreThanOneContent(otherColumnContent))
                bitEdition = utilities.getBitEdition(otherColumnContent);

            string[] otherComponents = utilities.separateAndOccurrences(otherColumnContent);
            foreach (string component in otherComponents)
            {
                if ((applications.applicationExist(component.Trim() + bitEdition)) && (!applications.applicationExist(originalContent.Trim() + bitEdition)))                                  
                    return true;                
            }
            return false;
       }
        
    }


     
}
