using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestMatrixApp
{
    class ErrorMessages
    {
        public ErrorMessages() { 
        }

        public void showErrorFileNotExist(string file){
            System.Windows.Forms.MessageBox.Show("File that contains the valid headers does not exit \n file: " + file,
                                            "ERROR: File does not exist",
                                            System.Windows.Forms.MessageBoxButtons.OK,
                                            System.Windows.Forms.MessageBoxIcon.Error,
                                            System.Windows.Forms.MessageBoxDefaultButton.Button1);                        
        }

        public void showErrorSheetNotExist(string file, string sheet)
        {
            System.Windows.Forms.MessageBox.Show("Sheet with the name '" + sheet + "' \n could be found in the Config TM file '" + file + "' \n Application is going to close ...  ",
                                            "ERROR: Sheet does not exist",
                                            System.Windows.Forms.MessageBoxButtons.OK,
                                            System.Windows.Forms.MessageBoxIcon.Error,
                                            System.Windows.Forms.MessageBoxDefaultButton.Button1);
            System.Environment.Exit(0);   
        }
        




        public void showError(string header, string message)
        {
            System.Windows.Forms.MessageBox.Show(message,
                                            header,
                                            System.Windows.Forms.MessageBoxButtons.OK,
                                            System.Windows.Forms.MessageBoxIcon.Error,
                                            System.Windows.Forms.MessageBoxDefaultButton.Button1);
        }
        
    }
}
