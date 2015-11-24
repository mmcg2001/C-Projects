using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SandG
{
    public partial class Form1 : Form
    {
        int n;
        int o;
        int r;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox1.Text == "")
            {
                n = 0;
            }
            else if (textBox2.Text == "")
            {
                o = 0;
            }
            else
            {
                n = Convert.ToInt16(textBox1.Text);
                o = Convert.ToInt16(textBox2.Text);
                r = n - o;
            }

            checkRemainder();
        }
        private void checkRemainder()
        {
            if (r % 2 == 0)
            {
                label1.Text = r.ToString();
            }
            else
            {
                label1.Text = "Not Divisible By 2";
            }
        }
    }
}
