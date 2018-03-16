/*
 *Author: Mamish Roy 
 *Purpose: Read Excel File and converts the cell values to SQL Script 
 *version : v1.01
 *Comments: Excell shal be preformated, as first row of the sheet is name of the coulmsn in the tables , sheet name should be Table name  
 * */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SQL_SCRIPT_EXTENSION
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ScriptGenerator());
        }
    }
}
