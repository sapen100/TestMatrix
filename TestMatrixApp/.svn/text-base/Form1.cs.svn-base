using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic;
using System.Reflection;
using System.Text.RegularExpressions;

namespace TestMatrixApp
{
    public partial class Form1 : Form
    {        
        static string   currentDir       = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
        static string   resultFolder     ;   
        static string   bulletinsFolder  ;        
        //string          serverDir        = @"\\pumpkin\Common Tools\PT Bulletins";
        //string          serverDir        = @"D:\T&A\BulletinsTM\Jul 2013 PT\all";
        string          serverDir = @"C:/";
        string          urlsBegin        = "http://technet.microsoft.com/en-us/security/bulletin/";
        string          workingDirectory;        
        TestMatrix      tm;
        Utilities       utilities;        
        Stopwatch sw = new Stopwatch();
        List<string> bulletins;
        public Form1()
        {            
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string resourceName = new AssemblyName(args.Name).Name + ".dll";
                string resource = Array.Find(this.GetType().Assembly.GetManifestResourceNames(), element => element.EndsWith(resourceName));

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
                {
                    Byte[] assemblyData = new Byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    return Assembly.Load(assemblyData);
                }
            };

            InitializeComponent();                                    
            tm        = new TestMatrix();            
            bulletins = new List<string>();   
         
        }

        private void button2_Click(object sender, EventArgs e)
        {            
            disableButtons();
            logMessage.Clear();
            logMessage.AppendText("EXECUTION STARTS...");
            if (prepareBulletin()){
                if ((rbWeb.Checked) && (!(textLastestTM.Text.Equals("")) && !validExcelFile(textLastestTM.Text)))   // verify the latest testMatrix file is correct                
                {
                    logMessage.AppendText("\n The latest TM to compare with field is not a valid exel file(*.xlsx). Please verify this path is correct");
                    return;
                }

                if (configCheckBox.Checked)
                    logMessage.AppendText("\n     Trying to download TMConfig.xlsx from SharePoint...");
                if (tm.defineconfigPath(configCheckBox.Checked))
                {
                    tm.setTMType(comboApp.SelectedItem.ToString(), rbLocal.Checked);
                    resetTimeWatch();
                    logMessage.AppendText("\n     START TRAVERSING ALL BULLETINS");
                    if (rbLocal.Checked)
                        traverseLocalBulletins();
                    else
                        traverseWebBulletins();

                    if (tm.endDocuments())
                    {
                        if ((rbWeb.Checked) && (!textLastestTM.Text.Equals("")))//There is an old matrix to compare with
                        {
                            logMessage.AppendText("\n Comparing the differences that exist between the old and new TM...\n");
                            tm.compareTMs(textLastestTM.Text, currentDir + "\\TMResult.xlsx");
                        }
                        tm.openResult();
                        this.Close();
                        return;
                    }

                    string ExecutionTimeTaken = string.Format("{0}:{1}:{2}", sw.Elapsed.Minutes, sw.Elapsed.Seconds, sw.Elapsed.TotalMilliseconds);
                    logMessage.AppendText("\nExecution end in " + ExecutionTimeTaken);
                }                                                
            }                       
            logMessage.AppendText("\nEXECUTION STOP...\n");
            enableButtons();
        }


        private void traverseLocalBulletins()
        {
            string[] fileEntries = Directory.GetFiles(bulletinsFolder);
            foreach (string fileName in fileEntries)
            {                    
                if ((fileName.ToUpper().Contains(".DOCX")) || (fileName.ToUpper().Contains(".DOC")))
                {
                    logMessage.AppendText("\n Processing Bulletin: " + fileName);
                    tm.processWordDocument(fileName);
                }
                else
                {
                    MessageBox.Show("ERROR: A *.docx file was expected, " + fileName + "is not a word document, Please verify and run it Again");
                    return;
                }
            }
        }

        private void traverseWebBulletins()
        {
            bulletins = formURLs(bulletins);
            if (textLastestTM.Text.Equals(""))  // if there is no comparisson to be made 
            {
                foreach (string bulletin in bulletins)
                {
                    logMessage.AppendText("\n Processing Bulletin: " + bulletin);
                    tm.processWebDocument(bulletin);
                }
            }
            else 
            {
                foreach (string bulletin in bulletins)
                {
                    logMessage.AppendText("\n downloading Bulletin: " + bulletin);
                    if (!tm.downloadWebDocument(bulletin))
                        logMessage.AppendText("\n We could not download Bulletin: " + bulletin);
                }
                logMessage.AppendText("\n TRAVERSING downloaded Bulletins");
                tm.processWebDocuments(textLastestTM.Text);
                    
            }

        }


        public List<string> formURLs(List<string> bulletins)
        {
            List<string> aux = new List<string>();
            foreach (string bulletin in bulletins)
                aux.Add(urlsBegin + bulletin);
            return aux;
        }

        public bool prepareBulletin()
        {
            if ( verifyBulletinInputText())
            {
                if (comboApp.SelectedItem != null)
                {                                    
                    bulletinsFolder = textFolderName.Text;
                    return true;                    
                }
                else
                {
                    logMessage.AppendText("\n You have to select Test Matrix Application you are trying to generate");
                    labelComoApp.ForeColor = Color.Red;
                    return false;
                }
            }
            else
            {
                logMessage.AppendText("\n Please verify the Bulletins Location. It seems not to be a valid folder");
                return false;
            }
        }

        public bool validExcelFile(string file)
        {
            if ((!file.Equals("")) && (File.Exists(file)) && (file.Trim().ToUpper().EndsWith(".XLSX")))
                return true;
            else
                return false;
        }

        private bool verifyBulletinInputText(){
            if (rbWeb.Checked)
                return verifyBulletinWebInputText();
            else
            {
                if ((!Directory.Exists(textFolderName.Text)) || (textFolderName.Text.Length < 3))
                    return false;
                return true;
            }
        }

        private bool verifyBulletinWebInputText()
        {
            bulletins.Clear();
            bool validateCorrelativity = false;
            string value = textBulletins.Text.Trim().ToUpper();
            if (value.Equals(""))
            {
                labelBulletins.ForeColor = Color.Red;
                logMessage.AppendText("You have to specify the bulletin(s) you want to download the patches from \n");
                return false;
            }

            if (value.Contains(","))
            {
                bulletins = value.Split(',').Select(item => item.Trim()).ToList();

            }
            else
            {
                if (value.Contains(" - "))
                {
                    bulletins = Regex.Split(value, " - ").Select(item => item.Trim()).ToList();
                    validateCorrelativity = true;
                }
                else
                    bulletins.Add(value);
            }

            return validateValues(validateCorrelativity);
        }

        private bool validateValues(bool validateCorrelativity)
        {
            string pattern = @"\bMS\d{2}-\d{3}\b";
            string value;
            foreach (string valuef in bulletins)
            {
                value = valuef.Trim();
                if (!(Regex.IsMatch(value, pattern)) || (value.Length != 8))
                {
                    logMessage.AppendText("This value '" + value + "' is not a valid bulletin name. \n To generate the TestMatrix from a single bulletin: MS##-### \n To generate the TestMatrix from a range of bulletins MS##-### - MS##-### \n To generate the TestMatrix from specific bulletins: MS##-### , MS##-###, MS##-###");
                    return false;
                }
            }

            if (validateCorrelativity)
            {
                if (bulletins.Count() != 2)
                {
                    logMessage.AppendText("Only one range is allowed. \n e.g. MS11-011 - MS11-015\n");
                    return false;
                }

                if (!isSmallerThan(bulletins[0], bulletins[1]))
                {
                    return false;
                }
                fillMiddleValues();
            }
            return true;
        }

        // When we have a range of values, we have to fill the values in the middle. e.g. Range: ms11-10 - ms11-12  the output must be: ms11-010, ms11-011, ms11-012
        // We suppose there are only two values =  one range
        private void fillMiddleValues()
        {
            string prepo = bulletins[0].Split('-')[0];
            int value1 = Convert.ToInt32(bulletins[0].Split('-')[1]);
            int value2 = Convert.ToInt32(bulletins[1].Split('-')[1]);
            for (int i = value1 + 1; i < value2; i++)
            {
                insertToSecondToLast(prepo, i);
            }
        }


        private string formatToThreeDigit(int value)
        {
            if (value < 10)
                return "00" + value.ToString();
            if (value < 100)
                return "0" + value.ToString();
            return value.ToString();
        }

        private void insertToSecondToLast(string preposition, int value)
        {
            string svalue = preposition + "-" + formatToThreeDigit(value);
            bulletins.Insert(bulletins.Count() - 1, svalue);
        }

        // is value1 smaller than value2: e.g. value1 = ms11-010    11 =  year  010 = bulletin. Year must be equal in value1 and in value2
        private bool isSmallerThan(string value1, string value2)
        {
            bool result = true;
            int year1 = Convert.ToInt32(value1.Split('-')[0].Substring(2));
            int year2 = Convert.ToInt32(value2.Split('-')[0].Substring(2));

            if (year1 != year2)
            {
                logMessage.AppendText("The year of the bulletins to be downloaded must be equal. \n '" + value1.Split('-')[0] + "' is NOT EQUAL to '" + value2.Split('-')[0] + "' \n\n");
                result = false;
            }

            int bulletin1 = Convert.ToInt32(value1.Split('-')[1]);
            int bulletin2 = Convert.ToInt32(value2.Split('-')[1]);

            if (bulletin1 > bulletin2)
            {
                logMessage.AppendText("The Bulletin value(" + bulletin1 + ") must be smaller than the second Bulletin(" + bulletin2 + ") \n e.g. MS11-011 - MS11-015\n");
                result = false;
            }
            return result;

        }

        private string getFileName(string file) {
            return file.Substring(file.LastIndexOf("\\")+1);
        }
       
        private void buttonFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description  = "Please select the folder where the Bulletins are located ";
            folderBrowserDialog1.SelectedPath = textFolderName.Text;            
            folderBrowserDialog1.ShowNewFolderButton = false;            
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath != null)
            {
                textFolderName.Text = folderBrowserDialog1.SelectedPath;                
            }            
        }
        
        private void updateWorkingDirectory() {
            workingDirectory = textFolderName.Text.Substring(0, textFolderName.Text.LastIndexOf("\\") + 1);         
        }

        private void LogMessage_TextChanged(object sender, EventArgs e)
        {
            logMessage.ScrollToCaret();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {            
            tm = null;
        }

        public string askForPassword()
        {
            return "infected";
        }
        /*private bool createResultFolders()
        {            
            if (Directory.Exists(resultFolder))
            {                
                if (utilities.DeleteDirectory(resultFolder) == false) 
                    return false;                
            }            
            Directory.CreateDirectory(resultFolder);
            Directory.CreateDirectory(bulletinsFolder);
            return true;
        }*/


        public bool uncompressBulletins()
        {
            string bulletinRarFile = "";
            try
            {
                int count = 0;
                string[] directoryEntries = Directory.GetDirectories(textFolderName.Text);
                foreach (string directory in directoryEntries)
                {
                    logMessage.AppendText("\n   Extracting Bulletin from " + directory);
                    bulletinRarFile = getBulletinRarFile(directory);
                    if (!bulletinRarFile.Equals(""))
                    {
                        utilities.uncompressFile(bulletinRarFile, bulletinsFolder);
                        count++;
                    }
                }
                if (count == 0) {
                    MessageBox.Show("INFO: The folder does not contain any Bulletin. Please try again and select a valid Bulletin Path");
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR WHILE UNCOMPRESSING FILE: " + bulletinRarFile + " \n\n" + ex.Message + "\n\n The application is going to close.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                utilities.closeApplication();
            }

            return true;
        }

        private string getBulletinRarFile(string directory)
        {
            string[] fileEntries = Directory.GetFiles(directory, "*.zip");
            foreach (string file in fileEntries)
            {
                //MessageBox.Show(file);
                if (file.ToUpper().Contains("BULLETIN "))
                    return file;
            }
            return "";
        }

        private void disableButtons() {
            tmButton.Enabled         = false;
            buttonFolder.Enabled     = false;
            labelBulletins.ForeColor = Color.Black;
            labelComoApp.ForeColor   = Color.Black;
        }

        private void enableButtons()
        {
            tmButton.Enabled = true;
            buttonFolder.Enabled = true;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.configCheckBox, "http://2010.corp.mcafee.com/sites/avertqa/rc_coordination/Processes%20Library/Other/TestMatrixConfig/TMConfig.xlsx");
            textFolderName.Text = serverDir;

            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.textBulletins, "MS11-010 - MS11-018 \n MS11-012, MS11-014, MS11-017");

            System.Windows.Forms.ToolTip ToolTip3 = new System.Windows.Forms.ToolTip();
            ToolTip3.SetToolTip(this.textLastestTM, "If you would like to compare this TM, you have to point to the latest TM generated based on Beta Bulletins");            
        }

        private void resetTimeWatch(){
            sw.Reset();
            sw.Start();
        }

        private void textFolderName_TextChanged(object sender, EventArgs e)
        {

        }

        private void textFolderName_Click(object sender, EventArgs e)
        {
            buttonFolder_Click(sender,e);
        }

        private void bulletinUncompressorToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            String bulletinUncompressorPath = Path.GetDirectoryName(Application.ExecutablePath);                     
            Process bu = new Process();
            //MessageBox.Show(bulletinUncompressorPath + "\\calc.exe");
            bu.StartInfo.FileName = bulletinUncompressorPath + "\\BulletinUncompressor.exe";            
            bu.Start();
            bu.WaitForExit();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show( "Test Matrix Generartor 5.7 \n Part of Tools  Automation Team \n McAfee Labs . All rights reserved.",  "About Test Matrix Generator");
        }

        private void textFolderName_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (rbLocal.Checked)
            {
                logMessage.Location = new Point(logMessage.Location.X, 80);
                TApicture.Location = new Point(TApicture.Location.X, 360);
                tmButton.Location = new Point(tmButton.Location.X, 360);
                this.Height = this.Height - 50;

                textFolderName.Visible = true;
                buttonFolder.Visible = true;
                labelBulletin.Visible = true;

                textBulletins.Visible = false;
                labelBulletins.Visible = false;

                logMessage.Height = 275;
            }
        }

        private void rbWeb_CheckedChanged(object sender, EventArgs e)
        {
            if (rbWeb.Checked)
            {
                textFolderName.Visible = false;
                buttonFolder.Visible = false;
                labelBulletin.Visible = false;

                textBulletins.Visible = true;
                labelBulletins.Visible = true;

                TApicture.Location = new Point(TApicture.Location.X, 410);
                tmButton.Location = new Point(tmButton.Location.X, 410);
                this.Height = this.Height + 50;


                logMessage.Location = new Point(logMessage.Location.X, 105);
                logMessage.Height = 300;
                logMessage.Refresh();
            }
        }

        private void textBulletins_TextChanged(object sender, EventArgs e)
        {

        }

        private void buttonLatestTM_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Please select the latest TM generated based on Beta Bulletins";
            openFileDialog1.InitialDirectory = textLastestTM.Text;
            openFileDialog1.Filter = "excel files(*.xlsx)|*.xlsx";
            openFileDialog1.ShowHelp = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textLastestTM.Text = openFileDialog1.FileName.Replace("\\", "/");
            }            
        }

        private void button1_Click(object sender, EventArgs e)
        {        
        }

        private void button1_Click_1(object sender, EventArgs e)
        {            
        }

    }
}
