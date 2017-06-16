using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using Excel = Microsoft.Office.Interop.Excel;

namespace DrawBallot
{
    public partial class DataForm : Form
    {
        SQLiteConnection mConn;
        SQLiteDataAdapter mAdapter;
        DataTable mTable= new DataTable();
        int Mode = 0;
        public DataForm()
        {
           
            InitializeComponent();
            
        }

        private void DataForm_Load(object sender, EventArgs e)
        {
            string mDbPath = Application.StartupPath + "/DatabaseFinal.sqlite";
            mConn = new SQLiteConnection("Data Source=" + mDbPath);
            mConn.Open();
            mAdapter = new SQLiteDataAdapter("SELECT * FROM [PARTICIPANT]", mConn);
            mAdapter.Fill(mTable);

            new SQLiteCommandBuilder(mAdapter);
            dataGridView1.DataSource = mTable;
        }

        private void loadButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel File (*.xlsx)|*.xlsx|07-2003 Excel File (*.xls)|*.xls|all file (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                SQLiteCommand cmd = new SQLiteCommand("DELETE FROM PARTICIPANT;", mConn);
                cmd.ExecuteNonQuery();
                //MessageBox.Show(ofd.FileName);
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
                Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int RowCount = xlRange.Rows.Count;
                int ColCount = xlRange.Columns.Count;

                string tempID = null;
                string tempFirstName = null;
                string tempLastName = null;

                for (int i = 1;i<=RowCount;i++)
                {
                    tempID = xlRange.Cells[i, 1].Value2.ToString();
                    tempFirstName = xlRange.Cells[i, 2].Value2.ToString();
                    tempLastName = xlRange.Cells[i, 3].Value2.ToString();
                    SQLiteCommand sqlCmd = new SQLiteCommand("Insert into participant (ID,FIRSTNAME,LASTNAME) values (@ID,@FN,@LN);",mConn);
                    SQLiteParameter p = new SQLiteParameter("@ID", System.Data.DbType.String);
                    p.Value = tempID;
                    sqlCmd.Parameters.Add(p);
                    p = new SQLiteParameter("@FN", System.Data.DbType.String);
                    p.Value = tempFirstName;
                    sqlCmd.Parameters.Add(p);
                    p = new SQLiteParameter("@LN", System.Data.DbType.String);
                    p.Value = tempLastName;
                    sqlCmd.Parameters.Add(p);

                    //try
                    //{
                    sqlCmd.ExecuteNonQuery();
                    //}
                    //catch(Exception ex)
                    //{
                    // MessageBox.Show(ex.Message);
                    //}

                    mAdapter = new SQLiteDataAdapter("SELECT * FROM [PARTICIPANT]", mConn);
                    DataTable mTable = new DataTable(); // Don't forget initialize!
                    mAdapter.Fill(mTable);
                    new SQLiteCommandBuilder(mAdapter);
                    dataGridView1.DataSource = mTable;
                }
            }
        }

        private void DataForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (mAdapter == null)
                return;
            mAdapter.Update(mTable);
        }

        private void bMod_Click(object sender, EventArgs e)
        {
            IDBox.Visible = true;
            FNBox.Visible = true;
            LNBox.Visible = true;
            bAdd.Enabled = false;
            bDel.Enabled = false;
            bMod.Enabled = false;
            loadButton.Enabled = false;
            bCon.Visible = true;
            bCan.Visible = true;
            Mode = 2;

        }

        private void bAdd_Click(object sender, EventArgs e)
        {
            IDBox.Visible = true;
            FNBox.Visible = true;
            LNBox.Visible = true;
            bAdd.Enabled = false;
            bDel.Enabled = false;
            bMod.Enabled = false;
            loadButton.Enabled = false;
            bCon.Visible = true;
            bCan.Visible = true;
            Mode = 1;
        }

        private void bCon_Click(object sender, EventArgs e)
        {
            if (Mode == 1)
            {
                var newrow = mTable.NewRow();
                newrow["ID"] = IDBox.Text;
                newrow["FIRSTNAME"] = FNBox.Text;
                newrow["LASTNAME"] = LNBox.Text;
                mTable.Rows.Add(newrow);
            }
            else
            {
                int selectedRow;
                selectedRow = dataGridView1.CurrentCell.RowIndex;
                DataGridViewRow row = dataGridView1.Rows[selectedRow];
                row.Cells[1].Value = IDBox.Text;
                row.Cells[2].Value = FNBox.Text;
                row.Cells[3].Value = LNBox.Text;

                dataGridView1.EndEdit();
                ((DataRowView)dataGridView1.CurrentRow.DataBoundItem).EndEdit();
            }

            IDBox.Visible = false;
            FNBox.Visible = false;
            LNBox.Visible = false;
            bAdd.Enabled = true;
            bDel.Enabled = true;
            bMod.Enabled = true;
            loadButton.Enabled = true;
            bCon.Visible = false;
            bCan.Visible = false;
            

            
        }

        private void bCan_Click(object sender, EventArgs e)
        {
            IDBox.Visible = false;
            FNBox.Visible = false;
            LNBox.Visible = false;
            bAdd.Enabled = true;
            bDel.Enabled = true;
            bMod.Enabled = true;
            loadButton.Enabled = true;
            bCon.Visible = false;
            bCan.Visible = false;
            Mode = 0;
        }

        private void bDel_Click(object sender, EventArgs e)
        {
            int selectedRow;
            selectedRow = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(selectedRow);
        }

        private void bReset_Click(object sender, EventArgs e)
        {
            mTable.Clear();
        }
    }
}
