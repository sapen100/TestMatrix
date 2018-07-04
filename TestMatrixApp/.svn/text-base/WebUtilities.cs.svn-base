using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HtmlAgilityPack;

namespace TestMatrixApp
{
    class WebUtilities
    {
        public WebUtilities()
        {
            
        }

        public string getTagContent(HtmlDocument page, string tagName)
        {
            try
            {
                if (page.DocumentNode != null)
                {
                    if (page.DocumentNode.SelectNodes("//" + tagName) != null)
                    {
                        return page.DocumentNode.SelectSingleNode("//" + tagName).InnerText.ToString().Trim();
                    }

                }
            }
            catch (NullReferenceException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                return "";
            }            
            return "";
            
            
        }
    }
}
