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
using System.Media;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using System.Windows.Forms;

namespace DrawBallot
{
    public partial class MainForm : Form
    {
        DataForm dataForm = new DataForm();
        //SQLiteConnection mConn;
        //SQLiteDataAdapter mAdapter;
        //DataTable mTable = new DataTable();
        List<string> IDList = new List<string>();
        int index;
        bool drawPushed = false;
        Timer timer = new Timer();
        
        Setting setting = new Setting();
        Winner winner = new Winner();
        public MainForm()
        {
            InitializeComponent();
            


        }

        private void drawButton_Click(object sender, EventArgs e)
        {
            Stream str = Properties.Resources.DrumRoll;
            SoundPlayer player = new SoundPlayer(str);
            player.Load();
            if (!drawPushed)
            {
                IDList.Clear();
                player.PlayLooping();
                //ADD id into IDlist
                int n = setting.Table.Rows.Count;
                string temp;
                for (int i = 0; i < n; i++)
                {
                    temp = setting.Table.Rows[i].Field<string>(0);
                    if (!setting.Table.Rows[i].Field<bool>(1))
                    {
                        IDList.Add(temp);
                    }
                }
                
                
                
                //}

                //
                IDList.Shuffle();
                index = 0;
                
                //

                timer.Interval = 150;
                timer.Tick += new EventHandler(timer_Tick);
                timer.Start();
                drawButton.Text = "Stop";
                drawPushed = true;
            }
            else
            {
                player.Stop();
                timer.Stop();
                drawButton.Text = "Draw";
                drawPushed = false;
                int n = cbQuantity.SelectedIndex + 1;
                DataRow result = winner.Table.NewRow();
                string temp;
                foreach (ListViewItem item in listView1.Items)
                {
                    result["ID"] = item.SubItems[0].Text;
                    result["First Name"] = item.SubItems[1].Text;
                    result["Last Name"] = item.SubItems[2].Text;
                    result["Prize"] = cbPrize.Text;
                    result["Date"] = DateTime.Now.ToString();
                    winner.Table.Rows.Add(result.ItemArray);
                    temp = IDList.Find(x => x.StartsWith(item.SubItems[0].Text));
                    DataRow row = setting.Table.Select("ID = '" + temp + "'").FirstOrDefault();
                    row[1] = true;
                }

                
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dataButton_Click(object sender, EventArgs e)
        {

        }

        void timer_Tick(Object Sender, EventArgs e)
        {
            int count = 0;
            DataRow result1 = setting.Table.NewRow();
            string[] arr = new string[3];
            List<string> arrL = new List<string>();
            foreach (ListViewItem item in listView1.Items)
            {
                arr = IDList[(index + count) % IDList.Count].Split('-');
                arrL = arr.ToList<string>();
                switch (arrL.Count)
                {
                    case 1:
                        arrL.Add("");
                        arrL.Add("");
                        break;
                    case 2:
                        arrL.Add("");
                        break;
                    default:
                        break;
                }

                
                item.SubItems[0].Text = arrL[0];
                item.SubItems[1].Text = arrL[1];
                item.SubItems[2].Text = arrL[2];
                count++;
                index++;
                if (index > IDList.Count)
                {
                    index = 0;
                }
            }
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        
        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            int n = 0;
            string[] arr = new string[3];
            
            for (int i = 0; i < cbQuantity.SelectedIndex + 1; i++)
            {
                arr[0] = "ID" + i.ToString(); ;
                arr[1] = "First Name";
                arr[2] = "Last Name";
                 
                listView1.Items.Add(Methods.CreateItem(arr));
                n++;
            }

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            dataForm.ShowDialog();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            setting.ShowDialog();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            
            cbQuantity.SelectedIndex = 0;
            //
            
            cbPrize.DataSource = setting.listPrize;
            //
            ToolStripControlHost host = new ToolStripControlHost(drawButton);
            ToolStripControlHost box1 = new ToolStripControlHost(cbQuantity);
            ToolStripControlHost box2 = new ToolStripControlHost(cbPrize);
            toolStrip1.Items.Add(host);
            toolStrip1.Items.Add(box1);
            toolStrip1.Items.Add(box2);
            //load exel file
            
            
            listView1.Columns.Add("ID", 40);
            listView1.Columns.Add("First Name", 200);
            listView1.Columns.Add("Last Name", 200);
        }

        private void toolStripButton2_Click_1(object sender, EventArgs e)
        {
            //history.ShowDialog();
        }

        private void optionsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            setting.ShowDialog();
        }


        
        public Image pBox1Image
        {
            get
            {
                return this.pictureBox1.Image;
            }
            set
            {
                this.pictureBox1.Image = value;
            }
        }

        private void winnerListToolStripMenuItem_Click(object sender, EventArgs e)
        {
            winner.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void listView1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            File.WriteAllText("Participant.csv", Methods.DataTableToCSV(setting.Table,';').ToString());
            File.WriteAllText("Winner.csv", Methods.DataTableToCSV(winner.Table, ';').ToString());
        }

        private void listView1_SelectedIndexChanged_2(object sender, EventArgs e)
        {

        }
    }
}
