/*Topic: creat insert/Update statement from the excel file. 
 * Authore: roym
 * Date: 20/03/2016
 * Refrence: EP Plus from souceforge 
 * Excel Format should be in liue with following rule
 * first row should consist of the coulmns name of the table.
 * table name should be the name of the sheet. 
 * for update statement  directly write or use the add button  to create where statement 
 *Rev0.1 - bug fixed first Row was skipping Date 13.08.2016
 *Rev0.2 - Enhanced by adding drop down for Sheet names
 * * */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace SQL_SCRIPT_EXTENSION
{
    public partial class ScriptGenerator : Form
    {
        public string FilePath { get; set; }
        public List<string> CoulmnName;
        public String TableName = string.Empty;
        public BindingList<string> Shetnames;

        public bool IsFileLoaded;

        ContextMenu Cm;
        
        public ScriptGenerator()
        {
            InitializeComponent();

            CoulmnName = new List<string>();
            MenuItem mnuCopy = new MenuItem("Copy");
            mnuCopy.Click +=new EventHandler(mnuCopy_Click);
            Cm = new ContextMenu();
            Cm.MenuItems.Add(mnuCopy);
            listBox1.ContextMenu = Cm;
            ButtonState(TableName);
            Shetnames = new BindingList<string>();
            lukedSheet.Properties.DataSource = Shetnames;
        }
        /// <summary>
        /// Right click to copy selected statement
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void mnuCopy_Click(object sender, EventArgs e)
        {
            string s1 = "";
            if (listBox1.Items.Count > 0)
            {
                string item = listBox1.SelectedItem.ToString() ?? string.Empty;
                s1 += item.ToString() + "\r\n";
                Clipboard.SetText(s1);
                MessageBox.Show("Statment Copied to ClipBord!!!!!");
            }
        }
        /// <summary>
        /// Load the Excel file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog.Multiselect = false;
            var Result = openFileDialog.ShowDialog();
            if (Result == DialogResult.OK)
            {
                FilePath = openFileDialog.FileName;
                tePath.Text = FilePath;
                if (!validateExtension())
                {
                    MessageBox.Show("Only Excel File is to be excecuted!!");
                }
                else
                {
                    IsFileLoaded = true;
                }

                using (var Package = new OfficeOpenXml.ExcelPackage())
                {
                    try
                    {
                        TableName = String.Empty;
                        using (var stream = File.OpenRead(FilePath))
                        {
                            Package.Load(stream);
                        }

                        // May Be asked here about the Sheet Number 
                        foreach (var item in Package.Workbook.Worksheets.ToList())
                        {
                            Shetnames.Add(item.ToString());
                        }

                       lukedSheet.Properties.DataSource = Shetnames;
                    }
                    finally
                    {
                        //TableName = String.Empty;
                    }
                    }
                
            }

        }
        /// <summary>
        /// Validate File Extension
        /// </summary>
        /// <returns></returns>

        private bool validateExtension()
        {
            var allowedExtension = new[] { ".xlsx",".xls",".xlsm" };

            if (allowedExtension.Contains(Path.GetExtension(FilePath)))
            {
                return true;
            }
            else
                return false;

            
        }
        /// <summary>
        /// Insert Statment Button Funcationality
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btSqlInsert_Click(object sender, EventArgs e)
        {
          
            if (IsFileLoaded)
            {
                List<string> Datasource = new List<string>();

                if (Datasource.Any())
                {
                    listBox1.Items.Clear();
                    Datasource.Clear();
                }
                Datasource = ConstructInsertStatement(GetMainTableFromFile());

                listBox1.DataSource = Datasource.ToList();
                groupBox2.BackColor = Color.Green;
            }
            else
                MessageBox.Show("Load A excel File");
            
        }
        /// <summary>
        /// Insert satement Creation
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public List<string> ConstructInsertStatement(DataTable table)
        {
            List<string> returnList = new List<string>();
            StringBuilder S = new StringBuilder();
           // roym Bug fixe intilize it from zero
            for (int i = 0; i < table.Rows.Count; i++)
            {
                DataRow NewRow = table.NewRow();
                S.Append("insert into ");
                S.Append(TableName);
                S.Append(" ( ");
                int Counter = 0;
                //Coulmn Names
                foreach (var item in CoulmnName)
                {
                    Counter++;
                    S.Append(item);
                    if (Counter == CoulmnName.Count)
                    {
                    }
                    else
                    S.Append(",");
                }

                S.Append(" ) values(");

                NewRow = table.Rows[i];
                int Newcounter = 0;
                // Row item 
                foreach (var item in CoulmnName)
                {
                    Newcounter++;
                    S.Append(" '");
                    S.Append(NewRow[item]);
                    if (Counter == Newcounter)
                    {
                        S.Append("'");
                    }
                    else
                    S.Append("' ,");
                }

                S.Append(");");

                returnList.Add(S.ToString());
                S.Clear();
            }
            return returnList;
        }
        /// <summary>
        /// Convert the excel fiel into data table 
        /// </summary>
        /// <returns></returns>
        public DataTable GetMainTableFromFile()
        {
            DataTable Ntable = new DataTable();
            CoulmnName.Clear();
            using (var Package = new OfficeOpenXml.ExcelPackage())
            {
                try
                {

                    using (var stream = File.OpenRead(FilePath))
                    {
                        Package.Load(stream);
                    }

                    // May Be asked here about the Sheet Number 
                    var WorkSheet = Package.Workbook.Worksheets[TableName];

                    //TableName = Package.Workbook.Worksheets.First().ToString();
                    // Extract the coulmns Names 
                    foreach (var Cell in WorkSheet.Cells[1, 1, 1, WorkSheet.Dimension.End.Column])
                    {
                        Ntable.Columns.Add(Cell.Value.ToString());
                        CoulmnName.Add(Cell.Value.ToString());
                    }

                    // extract the rows 
                    long CoulmnCounter = Ntable.Columns.Count;
                    object[] array = new object[CoulmnCounter];
                    int counter = 0;

                    for (int rowNum = 2; rowNum <= WorkSheet.Dimension.End.Row; rowNum++)
                    {
                        foreach (var item in WorkSheet.Cells[rowNum, 1, rowNum, WorkSheet.Dimension.End.Row])
                        {
                            array[counter] = item.Value;
                            counter++;
                        }
                        Ntable.Rows.Add(array);
                        counter = 0;
                        Array.Clear(array, 0, array.Length);
                    }
                }
                catch (FileNotFoundException ex)
                {
                    MessageBox.Show(string.Format("Could not Load File {0}. Please try Opening it Again!!!", ex.FileName));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Could not Load File {0}. Please try Opening it Again or formate of excel is wrong!!!", ex.Message));
                }

            }

            return Ntable;
        }
        /// <summary>
        /// Copy scripit to clipboard & Export to  local folder 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btCopy_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0)
            {
                string s1 = "";
                foreach (object item in listBox1.Items) s1 += item.ToString() + "\r\n";
                File.WriteAllText(Path.ChangeExtension(FilePath,".sql"), s1);
                Clipboard.SetText(s1);
                MessageBox.Show("All Statments are Copied to ClipBord File Created -->"+ Path.GetFileNameWithoutExtension(FilePath) +" Exported to Sorce Directory!!!!!");
            }
        }
        /// <summary>
        /// Update Statement Button CLick 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btSqlUpdate_Click(object sender, EventArgs e)
        {
            if (IsFileLoaded)
            {
                DataTable ReturnTable = new DataTable();
                ReturnTable = GetMainTableFromFile();
                using (var WhereCondition = new WhereForm(CoulmnName))
                {

                    WhereCondition.ShowDialog();
                    if (WhereCondition.DialogResult == DialogResult.OK)
                    {
                        
                        if (IsFileLoaded)
                        {
                            List<string> Datasource = new List<string>();

                            if (Datasource.Any())
                            {
                                listBox1.Items.Clear();
                                Datasource.Clear();
                            }
                            Datasource = CreateUpdateStatement(ReturnTable, WhereCondition.WhereStaement);

                            listBox1.DataSource = Datasource.ToList();
                            groupBox2.BackColor = Color.Green;
                        }
                    }
                }
            }
            else
                MessageBox.Show("Load A excel File");
        }
        /// <summary>
        /// Update statement 
        /// </summary>
        /// <param name="table"></param>
        /// <param name="WhereStaement"></param>
        /// <returns></returns>
        private List<string> CreateUpdateStatement(DataTable table, string WhereStaement)
        {
            List<string> returnList = new List<string>();
            StringBuilder S = new StringBuilder();

            for (int i = 1; i < table.Rows.Count; i++)
            {
                DataRow NewRow = table.NewRow();
                S.Append("Update ");
                S.Append(TableName);
                S.Append(" Set ");
                int Counter = 0;
                NewRow = table.Rows[i];
                foreach (var item in CoulmnName)
                {
                    Counter++;
                    S.Append(item);
                    if (Counter == CoulmnName.Count)
                    {
                        S.Append(" = '");
                        S.Append(NewRow[item]);
                        S.Append("' ");
                    }
                    else
                    {
                        S.Append(" = '");
                        S.Append(NewRow[item]);
                        S.Append("',  ");
                    }
                }
                S.Append(WhereStaement);
                S.Append(" ;");

                returnList.Add(S.ToString());
                S.Clear();
            }

            return returnList;
        }
        /// <summary>
        /// Reset 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btReset_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count > 0)
            {
                listBox1.DataSource = null;
            }
            tePath.Text = string.Empty;
            Clipboard.Clear();
            IsFileLoaded = false;
            CoulmnName.Clear();
            groupBox2.BackColor = Color.Gray;

            TableName = String.Empty;

            ButtonState(TableName);
            Shetnames.Clear();
        }

        private void btimpdb_Click(object sender, EventArgs e)
        {
           
        }

        private void lukedSheet_EditValueChanged(object sender, EventArgs e)
        {
            TableName = lukedSheet.Text.ToString();
            ButtonState(TableName);
        }

        private void ButtonState(string Sheetname)
        {
            if (String.IsNullOrEmpty(TableName))
            {
                btCopy.Enabled = false;
                btReset.Enabled = false;
                btSqlInsert.Enabled = false;
                btSqlUpdate.Enabled = false;
                btimpdb.Enabled = false;
            }
            else
            {
                btCopy.Enabled = true;
                btReset.Enabled = true;
                btSqlInsert.Enabled = true;
                btSqlUpdate.Enabled = true;
                btimpdb.Enabled = true;
            }
        }
    }
}
