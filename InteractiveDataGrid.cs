using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.DataAccess.Client; // ODP.NET Oracle managed provider
using Oracle.DataAccess.Types;
using System.Data.OleDb;


namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        //removed database connection info
		
        String mySelectQuery;
        OleDbConnection myConnection;
        OleDbCommand myCommand;
        DataTable data;
        OleDbDataAdapter da;


        public Form1()
        {
            InitializeComponent();
        }

      /*  private void button1_Click(object sender, EventArgs e)
        {
            
            OracleConnection conn = new OracleConnection();
            conn.ConnectionString = oradb; // C#
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "select team_id from teams where ARENA_ID = '"+ arenaTextBox.Text.ToUpper() +"'";
            cmd.CommandType = CommandType.Text;
            OracleDataReader dr = cmd.ExecuteReader();
            dr.Read();
            label1.Text = dr.GetString(0);
            conn.Dispose();
        }*/
           
        private void Form1_Load(object sender, EventArgs e)
        {

            mySelectQuery = "select p.player_id, p.p_lname, p.p_fname, p.team_id, p.position, s.games_played, s.goals, s.assists, (s.goals + s.assists)POINTS from players p, player_stats s where p.position != 'G' and p.team_id =s.team_id and p.player_id = s.player_id";

            myConnection = new OleDbConnection(sConnectionString);
            myCommand = new OleDbCommand(mySelectQuery, myConnection);

         
            myConnection.Open();
      
            da = new OleDbDataAdapter(myCommand);
           
            data = new DataTable();

            da.Fill(data);
            dataGridView1.DataSource = data;

            //generates A-Z in the ComboBox
            for(int i = 65; i <= 90; i++){
                comboBox1.Items.Add(char.ConvertFromUtf32(i));
            }


        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            myConnection.Close();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //contains the owningrow property needed
            DataGridViewCell cell = null;
            foreach (DataGridViewCell selectedCell in dataGridView1.SelectedCells)
            {
                cell = selectedCell;
                break;
            }
            if (cell != null)
            {
                //OwningRow gets the row that contains this cell
                //DataGridViewRow allows us to used the .Cells[] functionality
                DataGridViewRow row = cell.OwningRow;
                string newQuery = "Select p.team_ID, c.t_record, t.conf_id from players p,curr_team_stats c, teams t where p_lname = '" + row.Cells["p_lname"].Value.ToString() + "' and p.team_id  = c.team_id and t.team_id = c.team_id";

                //creating a new command to run the second query, using the param
                myCommand = new OleDbCommand(newQuery, myConnection);
                OleDbDataReader myReader = myCommand.ExecuteReader();
                try
                {
                    myReader.Read();
                    textBox1.Text = myReader.GetString(0);
                    textBox2.Text = myReader.GetString(1);
                    textBox3.Text = myReader.GetString(2);
                }
                catch (InvalidOperationException)
                {
                    MessageBox.Show("Pick a cell with data inside", "Empty selection", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            mySelectQuery = "select p.player_id, p.p_lname, p.p_fname, p.team_id, p.position, s.games_played, s.goals, s.assists, (s.goals + s.assists)POINTS from players p, player_stats s where p.position != 'G' and p.team_id = s.team_id and p.player_id = s.player_id and p.p_lname LIKE '" + comboBox1.SelectedItem + "%'";

            myConnection = new OleDbConnection(sConnectionString);
            myCommand = new OleDbCommand(mySelectQuery, myConnection);


            myConnection.Open();

            da = new OleDbDataAdapter(myCommand);

            data = new DataTable();

            da.Fill(data);
            dataGridView1.DataSource = data;
        }

        //private void playerTextBox_TextChanged(object sender, EventArgs e)
        //{
        //    mySelectQuery = "select p.player_id, p.p_lname, p.p_fname, p.team_id, p.position, s.games_played, s.goals, s.assists, (s.goals + s.assists)POINTS from players p, player_stats s where p.position != 'G' and p.team_id = s.team_id and p.player_id = s.player_id and p.p_lname LIKE '" + playerTextBox.Text + "%'";

        //    myConnection = new OleDbConnection(sConnectionString);
        //    myCommand = new OleDbCommand(mySelectQuery, myConnection);


        //    myConnection.Open();

        //    da = new OleDbDataAdapter(myCommand);

        //    data = new DataTable();

        //    da.Fill(data);
        //    dataGridView1.DataSource = data;
        //}
    }
}
