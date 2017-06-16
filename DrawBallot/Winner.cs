using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace DrawBallot
{
    public partial class Winner : Form
    {
        DataTable mTable = new DataTable();
        public Winner()
        {
            InitializeComponent();
            
            if (File.Exists("Winner.csv"))
            {
                mTable = Methods.ConvertCSVtoDataTable2("Winner.csv");
            }
            else
            {

                mTable.Columns.Add("ID");
                mTable.Columns.Add("First Name");
                mTable.Columns.Add("Last Name");
                mTable.Columns.Add("Prize");
                mTable.Columns.Add("Date");

            }
            dataGridView1.DataSource = mTable;
            dataGridView1.Columns["Date"].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
        }
        public DataTable Table
        {
            get
            {
                return this.mTable;
            }
            set
            {
                this.mTable = value;
            }
        }

        private void bDel_Click(object sender, EventArgs e)
        {
            int selectedRow;
            selectedRow = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(selectedRow);
        }

        private void bReset_Click(object sender, EventArgs e)
        {
            mTable.Rows.Clear();
        }
    }
}
