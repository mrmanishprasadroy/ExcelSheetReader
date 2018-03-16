using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SQL_SCRIPT_EXTENSION.Database
{
    public partial class DatabaseAccess : Form
    {
        private DataTable LocalTable;
        public DatabaseAccess()
        {
            InitializeComponent();
        }

        public DatabaseAccess(DataTable Source)
        {
            InitializeComponent();
            LocalTable = new DataTable();
        }
    }
}
