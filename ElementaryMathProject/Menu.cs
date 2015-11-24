using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            (new Add()).Show();
        }

        private void subButton_Click(object sender, EventArgs e)
        {
            this.Hide();
            (new Subtract()).Show();
        }

        private void multButton_Click(object sender, EventArgs e)
        {
            this.Hide();
            (new Muliply()).Show();
        }

        private void divButton_Click(object sender, EventArgs e)
        {
            this.Hide();
            (new Divide()).Show();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
