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
    public partial class Add : Form
    {
        Random random = new Random();

        int num;
        int num2;

        int iAnswer;
        int answer;

        int ansCorrect;
        int ansMissed;
        int totalAttempt;
       
        public Add()
        {
            InitializeComponent();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.Text == "")
                {
                    MessageBox.Show("Must attempt to answer the question.", "Blank Field", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    try
                    {
                        iAnswer = Convert.ToInt16(textBox1.Text);
                        checkAnswer();
                        generateNumbers();
                        totalAttempt++;
                        textBox1.Clear();
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

        private void Add_Load(object sender, EventArgs e)
        {
            generateNumbers();
        }

        private void checkAnswer()
        {
            answer = num + num2;

            if (iAnswer == answer)
            {
                MessageBox.Show("Correct");
                ansCorrect++;
            }
            else
            {
                MessageBox.Show("Wrong");
                ansMissed++;
            }

        }

        private void generateNumbers()
        {
            num = random.Next(0, 9);
            num2 = random.Next(0, 9);

            label1.Text = num.ToString();
            label1.ForeColor = Color.Red;

            label3.Text = num2.ToString();
            label3.ForeColor = Color.Blue;
        }

        private void Add_FormClosing(object sender, FormClosingEventArgs e)
        {
            MessageBox.Show("Correct: " + ansCorrect + "\nMissed: " + ansMissed + "\nTotal Attempted: " + totalAttempt);
            (new Form1()).Show();
        }

    }
}
