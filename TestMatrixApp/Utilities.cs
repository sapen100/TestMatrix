using System.Windows.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Security;
using Ionic;
using System.Net;
using System.Text.RegularExpressions;
using HtmlAgilityPack;



namespace TestMatrixApp
{
    class Utilities
    {     
        private ErrorMessages em;

        public Utilities(ErrorMessages em){
            this.em = em;
        }
       
        public void uncompressFile(string file, string unpackDirectory, string password="")
        {            
            using (var zip = Ionic.Zip.ZipFile.Read(file))
            {                
                zip.Password = password;    
                zip.ExtractAll(unpackDirectory);                
            }
        }

        //This method copies the content of a folder to another folder
        public void copyFolder(string source, string target)
        {            
            DirectoryInfo diSource = new DirectoryInfo(source);
            DirectoryInfo diTarget = new DirectoryInfo(target);
            copyFolderRecursive(diSource, diTarget);
        }

        private void copyFolderRecursive(DirectoryInfo source, DirectoryInfo target)
        {
           // Check if the target directory exists, if not, create it.
            if (Directory.Exists(target.FullName) == false)
            {
                Directory.CreateDirectory(target.FullName);
            }

            // Copy each file into it's new directory.
            foreach (FileInfo fi in source.GetFiles())
            {                
                fi.CopyTo(Path.Combine(target.ToString(), fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                copyFolderRecursive(diSourceSubDir, nextTargetSubDir);
            }
        }

        public void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

            }
            finally
            {
                GC.Collect();
            }
        }

        public bool DeleteDirectory(string target_dir)
        {
            try
            {
                bool result = true;

                string[] files = Directory.GetFiles(target_dir);
                string[] dirs = Directory.GetDirectories(target_dir);

                foreach (string file in files)
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    File.Delete(file);
                }

                foreach (string dir in dirs)
                {
                    if (!DeleteDirectory(dir))
                        return false;
                }

                Directory.Delete(target_dir, false);

                return result;
            }
            catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show("DELETE DIRECTORY FAILED: "+ex.Message);
                return false;
            }
        }

        public bool existWeb(string url, ref HtmlAgilityPack.HtmlDocument page) 
        {            
            HtmlWeb webGet     = new HtmlWeb();
            int numberTries    = 15;
            try
            {                                
                int tryCounter = 1;                
                while ( (page = webGet.Load(url)).DocumentNode.InnerText.ToUpper().Contains("THE PAGE YOU REQUESTED CANNOT BE FOUND"))
                {
                    if (tryCounter == numberTries)
                    {
                        MessageBox.Show("EROR: After trying " + numberTries.ToString() + " times, we could not download the following URL: \n " + url);
                        return false;
                    }
                    numberTries++;
                }                
                return true;                
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("We could not access the page: " + url + " \n" + e.Message);
                return false;
            }
        }

        //This method extracts the bit edition of the content.
        public string getBitEdition(string content)
        {

            int iniPos = content.LastIndexOf("(");
            int endPos = content.LastIndexOf(")");
            if (iniPos != -1 && endPos != -1)
            {
                string parenthesisContent = content.Substring(iniPos, (endPos - iniPos) + 1);
                /*if (parenthesisContent.ToUpper().Contains("-BIT EDITIONS"))
                {*/
                    return " " + parenthesisContent;
                //}
            }
            return "";
        }
        
        //It gets the text from a Cell from in the Table
        public string getTextFromCell(List<Novacode.Paragraph> paragraphs)
        {
            string result = "";
            if (paragraphs.Count() == 1)
                return paragraphs[0].Text;

            foreach (Novacode.Paragraph paragraph in paragraphs) {
                if (paragraph.Text.Equals(""))  //we are assuming that if it has an empty paragraph, we assume that it's an enter = "\n"
                    result = result + "\n";
                else 
                {
                    if ( !(result.Equals("")) && (!paragraph.Text.StartsWith("\n")) )
                        result += "\n" + paragraph.Text;
                    else
                        result+= paragraph.Text;
                }                    
            }
                
            return result;
        }


        //This method is separating " and " ocurrences
        public string[] separateAndOccurrences(string content)
        {
            content = removeContentIn(content, "()"); // We do not take into account the contetn that is in parenthesys

            bool flagComma = false;
            if (content.Contains(", and ")) //if we have ", and " we are replacing it by |||, once it is done, the last for of the method is going to redo this replacement
            {
                content = content.Replace(", and ","|||");
                flagComma = true;
            }

            string[] result = content.Split(new string[] { " and " }, StringSplitOptions.RemoveEmptyEntries);
            
            if (flagComma) {
                for(int i=0; i < result.Count() ; i++) {
                    if (result[i].Contains("|||"))
                        result[i] = result[i].Replace("|||", ", and ");
                }
            }
            return result;
        }

        public bool moreThanOneContent(string content)
        {            
            /*if (content.Contains(", and ")) //if we have ", and " we are replacing it by |||, once it is done, the last for of the method is going to redo this replacement
            {
                content = content.Replace(", and ", "|||");                
            }
            
            string[] result = content.Split(new string[] { " and " }, StringSplitOptions.RemoveEmptyEntries);*/
            return separateAndOccurrences(content).Count() > 1;
        }


        public string extractComments(string content)
        {
            string result = "";
            int iniPos;            
            while ((iniPos = content.IndexOf("*")) != -1)
            {
                result += "*";
                content = content.Substring(0, iniPos) + content.Substring(iniPos + 1);
            }            

            int endPos;                        
            while ((iniPos = content.IndexOf("[")) != -1)
            {
                endPos = content.IndexOf("]");
                if ((iniPos != -1) && (endPos != -1))
                {
                    result += content.Substring(iniPos, (endPos - iniPos) + 1);
                    content = content.Substring(0, iniPos) + content.Substring(endPos + 1);

                }
            }                            

            return result;                        
        }


        private string getPassword()
        {
            string result       = "";
            Form prompt         = new Form();
            prompt.Width        = 300;
            prompt.Height       = 150;
            prompt.Text         = "Accessing Sharepoint";
            prompt.MinimizeBox  = false;
            prompt.MaximizeBox  = false;
            prompt.MinimumSize = new System.Drawing.Size(300, 150);
            prompt.MaximumSize = new System.Drawing.Size(300, 150);
            Label textLabelUser = new Label()   { Left = 30,    Top = 20, Width = 80,  Text = "User Name" };
            TextBox textBoxUser = new TextBox() { Left = 110,   Top = 20, Width = 120, Text = "Nai-Corp\\" + Environment.UserName, Enabled = true };
            Label textLabelPass = new Label()   { Left = 30,    Top = 45, Width = 80,  Text = "Password" };
            TextBox textBoxPass = new TextBox() { Left = 110,   Top = 45, Width = 120, UseSystemPasswordChar = true };
            Button cancel       = new Button()  { Left = 80,    Top = 70, Width = 70,  Text = "Cancel"};
            Button confirmation = new Button()  { Left = 160,   Top = 70, Width = 70,  Text = "OK"};
            prompt.Controls.Add(textBoxPass);
            confirmation.Click += (sender, e) => { result = textBoxPass.Text; prompt.Close(); };
            textBoxPass.KeyDown += (sender, e) => { if (e.KeyCode == Keys.Enter) {result = textBoxPass.Text;  prompt.Close();} };
            cancel.Click        += (sender, e) => { result = ""; prompt.Close(); };            
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(cancel);
            prompt.Controls.Add(textLabelUser);
            prompt.Controls.Add(textBoxUser);
            prompt.Controls.Add(textLabelPass);            
            prompt.ShowDialog();            
            return result;
        }

        public bool copyFileFromSharePoint(string sharePointURL, string currentDir){
            WebClient client = new WebClient();
            string password  = getPassword();            
            if (!password.Equals(""))
            {
                client.Credentials = new NetworkCredential(Environment.UserName, password);
                try
                {
                    client.DownloadFile(sharePointURL, currentDir);
                    client.Dispose();                    
                    return true;
                }
                catch (Exception ex)
                {
                    client.Dispose();
                    em.showError("Error Accessing Sharepoint", "Application failed trying to download: " + sharePointURL + "\n\n" + ex.Message + "\nPlease try again and make sure your userName and password are correct");             
                }
             }
            return false;
        }

        public bool thisCellcontainsConponents(string content)
        {
            //return content.Contains("\n(KB");
            string pattern = @"\w*\(\s?(KB)?\d+\s?\)";                        
            Match match = Regex.Match(content.Trim().ToUpper(), pattern);
            return (match.Success);            
        }


        public int obtainInitKB(string content)
        {
            //return content.Contains("\n(KB");
            string pattern = @"\(\s?(KB)?\d+\s?\)";                        
            Match match = Regex.Match(content.Trim().ToUpper(), pattern);
            if (match.Success)
                return match.Index;
            else
                return -1;
        }

        public string getOnlyNumericValue(string value) 
        {
            string pattern = @"\d+";    
            Match match = Regex.Match(value.Trim(), pattern);
            if (match.Success)
                return value.Substring(match.Index);
            else
                return "";
        }

        public string deleteComments(string content)
        {
         /*   content = deleteAsterisks(content.Trim());
            return deleteBrackets(content);*/
            string result = "";
            int iniPos;
            while ((iniPos = content.IndexOf("*")) != -1)
            {
                result += "*";
                content = content.Substring(0, iniPos) + content.Substring(iniPos + 1);
            }

            int endPos;
            while ((iniPos = content.IndexOf("[")) != -1)
            {
                endPos = content.IndexOf("]");
                if ((iniPos != -1) && (endPos != -1))
                {
                    result += content.Substring(iniPos, (endPos - iniPos) + 1);
                    content = content.Substring(0, iniPos) + content.Substring(endPos + 1);

                }
            }

            if (content.EndsWith("\n"))   //Delete if there is an enter at the end
                return content.Substring(0, content.IndexOf("\n"));
            
            return removeExtraSpaces(content.Trim());
        }

        public void closeApplication(){
            System.Environment.Exit(0);
        }
        public string deleteAsterisks(string content)
        {
            if (content.Contains("*"))
            {
                while (content.IndexOf("*") != -1)
                {
                    int iniPos = content.IndexOf("*");
                    if (iniPos != -1)
                        content = content.Substring(0, iniPos) + content.Substring(iniPos + 1);
                }

            }
            return content;
        }

        public string deleteBrackets(string content)
        {
            if (content.Contains("["))
            {
                int iniPos = content.IndexOf("[");
                int endPos = content.IndexOf("]");
                if ((iniPos != -1) && (endPos != -1))
                {
                    return content.Substring(0, iniPos) + content.Substring(endPos + 1);
                }
            }
            return content;
        }

        public string getBulletinReplacedBy(string content, string kb) 
        {
            string[] bulletinsRep = content.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
            kb = deleteComments(kb);
            string bulletinRep;
            foreach (string value in bulletinsRep)
            {
                bulletinRep = value.Trim().ToUpper();
                if (bulletinRep.EndsWith("REPLACED BY " + kb))
                    return value;
            }
            return content;
        }


        /*
         * This method deletes the parenthesis that exist between the name of the application and the KB
         * 1ST CASE
         *  Microsoft Office 2003 Service Pack 3
         *  (Windows common controls)
         *  (KB2597112)
         *  
         * 2nd CASE
         *  Microsoft Office 2003 Service Pack 3         
         *  (KB2597112)
         *  (Windows common controls)
        */
        public string[] deleteCommentsInTheMiddle(string[] components) 
        {
            List<string> result = new List<string>();
            int number = 1;
            string pattern = @"\(\s?(KB)?\d+\s?\)";    
            foreach (string component in components) 
            {                
                if ((number % 2) == 0)
                {
                    Match match = Regex.Match(component.Trim().ToUpper(), pattern);
                    if (match.Success)
                    {
                        result.Add(component);
                    }
                    else
                        number--;
                }
                else 
                {
                    if (!component.Trim().ToUpper().StartsWith("("))  //2ND CASE
                        result.Add(removeContentIn(component, "()"));
                }
                number++;
            }
            return result.ToArray();            
        }


        public string deleteKBContent(string value) {
            string pattern = @"\(\s?(KB)?\d+\s?\)";    
            Match match = Regex.Match(value.Trim().ToUpper(), pattern);
            if (match.Success)
                return value.Substring(0, match.Index);
            else
                return value;
        }

        public string removeExtraSpaces(string input)
        {
            string spaces = "  ";
            
            while (input.Contains("\n"))
            {
                input = input.Replace("\n", " ");
            }
            while (input.Contains(spaces))
            {
                input = input.Replace(spaces, " ");
                spaces = spaces + " ";
            }
            
            while (input.Contains(") ("))
            {
                input = input.Replace(") (", ")(");                
            }

            while (input.Contains(" ("))
            {
                input = input.Replace(" (", "(");
            }            

            return input.Trim();
        }

        /*This method deletes the content that is beteween the value "contentIn"
         * @contentin it's the values we are looking for to delete. This could be "()" or "[]"
         * */

        public string removeContentIn(string value, string contentIn){            
            int iniPos, endPos;
            string result = "";
            if ( (value.IndexOf(contentIn[0]) == -1) || !(numberCharsMatch(value, contentIn)) )
                return value;

            while ((iniPos = value.IndexOf(contentIn[0])) != -1) {                
                if ((endPos = value.IndexOf(contentIn[1])) != -1)
                {
                    result = value.Substring(0, iniPos) + value.Substring(endPos + 1); 
                    value  = value.Substring(0, iniPos) + value.Substring(endPos + 1);                    
                }
                else
                    return value;
            }
            return result.Trim();
        }

        private bool numberCharsMatch(string content, string contentIn) 
        {
            int countOpening = count(content, contentIn[0]);
            int countClosing = count(content, contentIn[1]);
            return (countOpening == countClosing) ? true : false;
                
        }

        private int count(string content, char character) { 
            int counter = 0;
            int initPos;
            while( (initPos = content.IndexOf(character)) != -1)
            {
                content = content.Substring(initPos + 1);
                counter++;
            }
            return counter;
        }

        //The content we are receving could be from bulletins or from the web(html), 
        //When the content comes from bulletins(doc), the content comes separated by "\n", when it comes from the web, it comes in a single line
        public string[] splitContent(string content)
        {
            if (content.Contains("\n"))     // it comes like this, when it comes from word bulletins
                content =  content.Replace("\n", " ");
                
                string pattern = @"\(\s?(KB)?\d+\s?\)";
                Match match;
                List<string> result = new List<string>();

                if (Regex.Match(content.Trim().ToUpper(), pattern).Index == 0) // we could not even found a parenthesis
                    return (new string[]{content});
                
                while ((match = Regex.Match(content.ToUpper(), pattern)).Index != 0)
                {
                    result.Add( content.Substring(0, match.Index));    // This is the content
                    content = content.Substring(match.Index);
                    result.Add(content.Substring(0, content.IndexOf(")") + 1)); // This is the KB
                    content = content.Substring(content.IndexOf(")") + 1).Trim();
                    if (content.Trim().StartsWith("(")) //If there is a comment at the end after the KB
                    {
                        result[result.Count - 2] = result[result.Count - 2].Trim() + " " + content.Substring(content.IndexOf("("), content.IndexOf(")") - content.IndexOf("(") + 1);
                        content = content.Substring(content.IndexOf(")") + 1);
                    }

                }
                return result.ToArray();            
        }

        public string deleteParenthesis(string str)
        {
            str = str.Trim();
            int iniPos = str.IndexOf("(");
            int endPos = str.IndexOf(")");
            if (iniPos != -1 && endPos != -1)
            {
                return str.Substring(iniPos + 1, (endPos - iniPos) - 1);
            }
            return str;
        }



        public string getBulletin(string input)
        {
            input = input.ToUpper();
            int inipos = input.IndexOf("MS");
            int endpos = input.IndexOf(" - ");
            if ((inipos > 0) && (endpos > 0))
                return input.Substring(inipos, (endpos - inipos) + 1).Trim();
            else
                return "";   // 
        }


        public string getDescription(string input)
        {
            input = input.ToUpper();
            //TODO: you gotta make a regex to see whethere what you are looking for is in the input
            //string pattern = "^MS//d{2}-d{*}$";
            //if (Regex.IsMatch(input, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))                        
            int endpos = input.IndexOf(" (");
            if (endpos > 0)         
                return input.Substring(0, endpos);
            else
                return "";   //             
        }


        public string getWebDescription(string input)
        {

            input = input.ToUpper();
            //TODO: you gotta make a regex to see whethere what you are looking for is in the input
            //string pattern = "^MS//d{2}-d{*}$";
            //if (Regex.IsMatch(input, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))                        
            int iniPos = input.IndexOf(" : ");
            int endpos = input.IndexOf(" (");            
            if ( (endpos > 0) && (iniPos > 0) )
                return input.Substring(iniPos + 3, (endpos - iniPos) - 3).Trim();
            else
                return "";   //             
        }

        public string getKB(string input)
        {
            int inipos = input.LastIndexOf(" (");
            int endpos = input.LastIndexOf(")");
            if ((inipos > 0) && (endpos > 0))
            {
                inipos = inipos + 2;               
                return input.Substring(inipos, (endpos - inipos));
            }
            else
                return "";   //             
        }       
    }
}
