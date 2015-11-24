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
    public partial class Divide : Form
    {
        Random random = new Random();

        int num;
        int num2;

        int iAnswer;
        int answer;

        int ansCorrect;
        int ansMissed;
        int totalAttempt;

        int remainder;
        int rAnswer;

        public Divide()
        {
            InitializeComponent();
        }
        
        private void Divide_Load(object sender, EventArgs e)
        {
            generateNumbers();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (num2 > num)
                {
                    answer = num2 / num;
                    remainder = num2 % num;
                }
                else
                {
                    answer = num / num2;
                    remainder = num % num2;
                }

                 if (textBox1.Text == "")
                {
                    MessageBox.Show("Must attempt to answer the question.", "Blank Field", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
              
                else
                {  
                    rAnswer = 0;
                    try
                    {
                        iAnswer = Convert.ToInt16(textBox1.Text);
                        DialogResult d = MessageBox.Show("Is there a remainder?", "Remainder", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (d == DialogResult.Yes)
                        {
                            groupBox1.Visible = true;
                            textBox2.Focus();
                        }
                        else
                        {
                            checkAnswer();
                            generateNumbers();
                            totalAttempt++;
                            textBox1.Clear();
                        }
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("Must input a number", "Wrong Input", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        textBox1.Clear();
                        textBox1.Focus();
                    }
                }
            }
        }
        private void checkAnswer()
        {
            if (iAnswer == answer && remainder == rAnswer)
            {
                MessageBox.Show("Correct");
                ansCorrect++;
            }
            else
            {
                MessageBox.Show("Wrong");
                ansMissed++;
            }
            textBox1.Focus();
        }

        private void generateNumbers()
        {
            num = random.Next(1, 10);
            num2 = random.Next(1, 10);

            if (num2 > num)
            {
                label1.Text = num2.ToString();
                label3.Text = num.ToString();
            }
            else
            {
                label1.Text = num.ToString();
                label3.Text = num2.ToString();
            }
            label1.ForeColor = Color.Red;
            label3.ForeColor = Color.Blue;
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox2.Text == "")
                {
                    MessageBox.Show("Must attempt to answer the question.", "Blank Field", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    try
                    {
                        rAnswer = Convert.ToInt16(textBox2.Text);
                        checkAnswer();
                        generateNumbers();
                        totalAttempt++;
                        textBox1.Clear();
                        textBox2.Clear();
                        groupBox1.Visible = false;
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("Must input a number", "Wrong Input", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                        textBox1.Clear();
                        textBox1.Focus();
                    }
                }
            }
        }

        private void Divide_FormClosing(object sender, FormClosingEventArgs e)
        {
            MessageBox.Show("Correct: " + ansCorrect + "\nMissed: " + ansMissed + "\nTotal Attempted: " + totalAttempt);
            (new Form1()).Show();
        }   
    }
}
