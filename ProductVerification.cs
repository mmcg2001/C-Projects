using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb; // must include for OleDb connection
using System.Net.Mail; // must include for mail
using System.Net; // must include to connect to the internet
using Outlook = Microsoft.Office.Interop.Outlook; // must include to send mail through outlook
using Excel = Microsoft.Office.Interop.Excel; // must include to write a file to excel
using System.Collections; // must include to ArrayLists
using System.IO; //must include for file input and stream readers
using System.Threading; // must include to use the File class and the Streamreader


//copy-copy-3
namespace ProductVerification
{
    public partial class ProductVerification : Form
    {
	
		//first project of the sort, used a local Access db for the matching
	
        //-----------CREATING THE GLOBAL VARIABLES FOR THE FORM-----------//

        //setting variable user equal to who logged in
        string user = Login.user;

        //instantiating the variable for the datatable
        DataTable Tbl = null;

        //initializing the variables for the numerical start and end times
        System.DateTime lVerfStart;
        System.DateTime lVerfEnd;
      

        //connection path to the db ... datasource removed
        String conn_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
 
        //creating the string for the query
        String sql = "";

        //string for the delivery number
        String deliveryNumber = "";

        //initalizing the OleDbConnection
        OleDbConnection conn = null;

        //initializing the strings for item and upc
        string item = "";
        string upc = "";

        //creating the integers for quantity and count
        int quantity = 0;
        int count = 1;
        int countResult;
        int quantityRemaining;
        int multiplier;

        //initializing the variables for the times
        string verificationStartTime = "";
        string verificationEndTime = "";
        string timeElapsed = "";


        //initializing the string for the excel file path
        string ExcelFilePath = "";

        //initalizing counts to decide to trigger the on close email or not
        int correctCount = 0;
        int shortageCount = 0;

        //-----------END OF VARIABLE CREATION-----------//

        public ProductVerification()
        {
            InitializeComponent();
        }

        //-----------ON FORM LOAD-----------//
        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                // set the user name into the userToolStripMenuItem
                userToolStripMenuItem.Text = user;
                //sets the conn varible to the conn_string to open the connection
                conn = new OleDbConnection(conn_string);
                //opens the connection
                conn.Open();


                //populating the drop down menu
                multiplierComboBox.Items.Add("1");
                multiplierComboBox.Items.Add("10");
                multiplierComboBox.Items.Add("20");
                multiplierComboBox.Items.Add("50");
                multiplierComboBox.Items.Add("100");

                //setting the default value for the combobox
                multiplierComboBox.SelectedIndex = 0;

                //creating a Datatable
                Tbl = new DataTable();
                //C reating the columns of the table
                Tbl.Columns.Add("Delivery Number");
                Tbl.Columns.Add("Item Number");
                Tbl.Columns.Add("UPC");
                Tbl.Columns.Add("Total Quantity");
                Tbl.Columns.Add("Quantity Scanned");
                Tbl.Columns.Add("Quantity Remaining");
                Tbl.Columns.Add("Start Time");
                Tbl.Columns.Add("End Time");
                Tbl.Columns.Add("Time Elapsed (mm:ss)");
            }
            catch (Exception ex)
            {
                //shows the error message
                MessageBox.Show(ex.Message);
            }
        }
        //-----------END OF ON FORM LOAD-----------//


        //-----------START OF runQuery METHOD-----------//

        //creating the functionality for the run_Query method
        private void run_Query()
        {
            //variables for user input
            item = itemTextBox.Text;
            upc = upcTextBox.Text;
            deliveryNumber = deliveryTextBox.Text;

            //checking to see which radio button is selected
            if (upcRadioButton.Checked == true)
            {
               //creating the query for standard upc
               sql = "SELECT * FROM upcTable WHERE Item = '" + item + "'";
            }
            else if (janRadioButton.Checked == true)
            {
                    //creating the query for jan upc
                    //q = select statement for jan upc
            }
            try
            {
                //OleDbCommand represents an SQL statement or stored procedure to execute against a data source
                //passing the SQL query and connection into the command to be run
                OleDbCommand cmd = new OleDbCommand(sql, conn);

                /*When we started to read from an OleDbDataReader it should always be open and positioned prior 
                  to the first record. The  Read() method in the OleDbDataReader is used to read the rows from 
                  the OleDbDataReader and it always moves forward to a new valid row, if any row exist. */
                OleDbDataReader reader = cmd.ExecuteReader();

                //Creating an ArrayList for the scanned items
                ArrayList scannedItems = new ArrayList();

                if (reader.HasRows == false)
                {
                    MessageBox.Show("No Records Found, please re-scan the item number. If problem persists please notify a supervisor", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    itemTextBox.Enabled = true;
                    itemTextBox.Clear();
                    itemTextBox.Focus();
                }
                else{
                //actions performed while the condition is true
                while (reader.Read())
                {
                    //Creating an array to store the fields variable
                    string[] fields = new string[1];

                    //setting the fields array variable equal to the upc generated by the sql statement
                    fields[0] = reader["UPC"].ToString();

                    //adding the field to the scanned item list
                    scannedItems.Add(fields);
                }

                //Looping through the ArrayList to verify the products
                foreach (string[] fields in scannedItems)
                {
                    //setting sql_upc equal to the field holding the UPC result from the query
                    string sql_upc = fields[0];

                   
                        //check to see if count is less than the quantity
                        //if the condition is true
                        if (quantity >= count)
                        {
                          
                            //checking if the user scan is equal to the queried result
                            if (upc.Equals(sql_upc))
                            {
                                //show the correct label, hide the supervisor label
                                correctLabel.Visible = true;
                                supervisorLabel.Visible = false;

                                //set the multiplier = 100 if 100 is selected in the combobox
                                if (multiplierComboBox.SelectedIndex == 4)
                                {
                                    multiplier = 100;
                                }
                                //set the multiplier = 50 if 50 is selected in the combobox
                                else if (multiplierComboBox.SelectedIndex == 3)
                                {
                                    multiplier = 50;
                                }
                                //else set the multiplier = 20 if 20 is selected in the combobox
                                else if (multiplierComboBox.SelectedIndex == 2)
                                {
                                    multiplier = 20;
                                }
                                //else set the multiplier = 10 if 10 is selected in the combobox
                                else if (multiplierComboBox.SelectedIndex == 1)
                                {
                                    multiplier = 10;
                                }
                                //else set the multiplier = 1 if 1 is selected in the combobox
                                else if (multiplierComboBox.SelectedIndex == 0)
                                {
                                    multiplier = 1;
                                }
                                //by default set the multiplier = 1
                                else
                                {
                                    multiplier = 1;
                                }

                                //increment the count by multiplier
                                count = count + multiplier;

                                //manuipuate data for desired output on label
                                countResult = count - 1;
                                countLabel.Text = countResult.ToString();

                                //calls the checkMultiplier method
                                checkMultiplier();

                                //setting the correctCount to initalize the closing methods
                                correctCount += 1;

                            }
                            //if condition is false
                            else
                            {
                                //show the see supervisor label, show error message.
                                supervisorLabel.Visible = true;
                                correctLabel.Visible = false;
                                MessageBox.Show("Item Number and UPC do not match!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                verificationEndTime = DateTime.Now.ToLongTimeString();

                                //setting the email subject and body
                                string subject = "Wrong Item Scan on order: " + deliveryNumber;
                                string body = "UPC Scanned: " + upc + " UPC Expected:" + sql_upc + " at: " + verificationEndTime + ". Scanned by: " + user;

                                //testing to see if continued, and a shortage is reported that that correct number is reported.
                                countResult -= 1;

                                //thread used to perform the sendEmail Method
                                Thread wEmail = new Thread(() => sendEmail(subject, body));
                                wEmail.IsBackground = true;
                                wEmail.Start();
                            }
                        }
                    }
                }
            }

            catch (Exception)
            {
                //displays the error message, sets the supervisor label to visible and the correct label to invisible
                MessageBox.Show("Item number and UPC do not match, please see supervisor!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                supervisorLabel.Visible = true;
                correctLabel.Visible = false;
            }
        }
        
        //-----------END OF runQuery METHOD-----------//


        //-----------START OF submitButton_Click event-----------//
        private void submitButton_Click(object sender, EventArgs e)
        {
            //check to see if the quantity is greater than the count
            //if condidtion is true
            if (quantity > countResult)
            {
                //calculating the shortage amount
                int result = quantity - countResult;

                //setting the end time
                verificationEndTime = DateTime.Now.ToLongTimeString();

                //setting the email subject and body
                string subject = "Shortage on order: " + deliveryNumber;
                string body = "Quantity expected: " + quantity.ToString() + ". Quantity scanned: " + countResult.ToString() +
                    ". This item is short by: " + result.ToString() + " at: "+ verificationEndTime +". Scanned by: " + user;
                
                //thread used to perform the sendEmail Method
                Thread sEmail = new Thread(() => sendEmail(subject, body));
                sEmail.IsBackground = true;
                sEmail.Start();

                //setting the end verification time to calculate duration
                lVerfEnd = DateTime.Now;
                //creating the variable to hold the duration
                TimeSpan time = lVerfEnd - lVerfStart;
                //creating the variable used to insert into excel
                timeElapsed = new DateTime(time.Ticks).ToString(@"mm:ss");

                //calls the cleaningUp method
                cleaningUp();

                //setting the shortageCount to initalize the closing methods
                shortageCount += 1;
            }
        }

        //-----------END OF submitButton_Click event-----------//

        
        //-----------START OF ON FORM CLOSE-----------//
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //if the form is closed sent back to the main hub
            (new MainHub()).Show();

            //check to see if running these methods is necessary
            if (correctCount > 0 || shortageCount > 0)
            {
                //thread used to perform the writeToExcel method
                //thead used to allow the UI to remain functioning
                Thread write = new Thread(writeToExcel);
                write.IsBackground = true;
                write.Start();
                //finishes this thread before starting the email thread
                write.Join();

                //subject title and body for email
                string subject = user +" Closing Excel Report";
                string body = "Excel report is attached";
                //sendEmail(subject, body);

                //thread used to perform the sendEmail Method
                Thread email = new Thread(() => sendEmail(subject, body));
                email.IsBackground = true;
                email.Start();
            }
                //closes the connection
                conn.Close();
        }
        //-----------END OF ON FORM CLOSE-----------//


        //-----------START OF sendEmail METHOD-----------//
        private void sendEmail(string subject, string body)
        {
            //intializes the array list
            ArrayList arEmails = new ArrayList();
            //checks to see if the file exists, to get all the recipients
            if (File.Exists("emailAddresses.txt") == true)
            {
                //setting a variable to read in each character from the file.
                string line;
                //set the file to the streamreader for the file
                var file = new System.IO.StreamReader("emailAddresses.txt");
                //loop thru the file until it hits it's end point, adds each item to the array list
                while ((line = file.ReadLine()) != null)
                {
                    arEmails.Add(line);
                }
                //closes the file
                file.Close();
            }

            //process for sending an email
            // Create the Outlook application.
            Outlook.Application oApp = new Outlook.Application();
            // Create a new mail item.
            Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
            // Set HTMLBody. 
            //add the body of the email
            oMsg.HTMLBody = body;
            //Subject line
            oMsg.Subject = subject;
            // Add a recipient.
            Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
            // Loop thru the arEmails ArrayList to sent emails to all recipients
            for (int i = 0; i < arEmails.Count; i++)
            {
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(arEmails[i].ToString());
                oRecip.Resolve();
            }
            if (ExcelFilePath != "")
            {
                oMsg.Attachments.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + ExcelFilePath, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1, "Report");
            }
            
            // Send
            oMsg.Send();
            // Clean up.
            oRecips = null;
            oMsg = null;
            oApp = null;
        }
        //-----------END OF sendEmail METHOD-----------//


        //-----------START OF resetValues METHOD-----------//
        private void resetValues()
        {
            //clearing the textboxes
            deliveryTextBox.Clear();
            itemTextBox.Clear();
            upcTextBox.Clear();
            quantityTextBox.Clear();

            //enabling textboxes
            deliveryTextBox.Enabled = true;
            itemTextBox.Enabled = true;
            quantityTextBox.Enabled = true;
            
            //setting the cursor back inside the delivery textbox
            deliveryTextBox.Focus();

            //resetting the labels and count variable
            countLabel.Text = "";
            count = 1;
            quantity = 0;
            quantityRemaining = 0;
            countResult = 0;

            //making both labels invisible
            correctLabel.Visible = false;
            supervisorLabel.Visible = false;

            //unselecting the radio buttons
            upcRadioButton.Checked = false;
            janRadioButton.Checked = false;

            //resetting the timers
            verificationStartTime = "";
            verificationEndTime = "";

            //making multiplier invisible
            groupBox1.Visible = false;

            //clearing items from combobox to eliminate duplicate values
            multiplierComboBox.Items.Clear();

            //reset the selections for the combobox
            multiplierComboBox.Items.Add("1");
            multiplierComboBox.Items.Add("10");
            multiplierComboBox.Items.Add("20");
            multiplierComboBox.Items.Add("50");
            multiplierComboBox.Items.Add("100");

            //resetting the default value for the combobox
            multiplierComboBox.SelectedIndex = 0;
        }
        //-----------END OF resetValues METHOD-----------//

        
        //-----------START OF resetButton_Click event-----------//
        private void resetButton_Click(object sender, EventArgs e)
        {
            //calling the resetValues method
            resetValues();
        } 
        //-----------END OF resetButton_Click event-----------//


        //-----------START OF KeyDown events-----------//

        //methods for each text box to perform when the enter key is pressed
        private void upcTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            //if the enter key is pressed or simulated run this block
            {
                if (e.KeyCode == Keys.Enter)
                {
                    //checks if the textbox is blank, if it is run this block
                    if (upcTextBox.Text == "")
                    {
                        MessageBox.Show("Enter UPC Number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        upcTextBox.Focus();
                    }
                    //if it is not blank, run this block
                    else
                    {
                        //calls the run_Query Method
                        run_Query();

                        //Clears the upcTextBox 
                        upcTextBox.Clear();
                    }
                }
            }
        }

        private void deliveryTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //checks if the textbox is blank, if it is run this block
                if (deliveryTextBox.Text == "")
                {
                    MessageBox.Show("Enter Delivery Number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    deliveryTextBox.Focus();
                }
                //if it is not blank, run this block
                else
                {
                    deliveryTextBox.Enabled = false;
                    
                    //start time initiated as a string to be exported into excel
                    verificationStartTime = DateTime.Now.ToLongTimeString();
                    //variable to hold the numerical start time used to calculate duration
                    lVerfStart = DateTime.Now;
                   
                    //setting the upcRadioButton as selected
                    upcRadioButton.Checked = true;

                    //Cursor placed in itemTextBox
                    itemTextBox.Focus();
                }
            }
        }

        private void itemTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //checks if the textbox is empty, if it is run this block
                if (itemTextBox.Text == "")
                {
                    MessageBox.Show("Enter Item Number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    itemTextBox.Focus();
                }
                //if it is not blank, run this block
                else
                {
                    //disable the itemTextBox, and put the focus in the quantityTextBox
                    itemTextBox.Enabled = false;
                    quantityTextBox.Focus();
                }
            }
        }

        private void quantityTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            //checks to see if the enter key was pressed
            if (e.KeyCode == Keys.Enter)
            {
                //checks if the textbox is empty, if it is run this block
                if (quantityTextBox.Text == "")
                {
                    MessageBox.Show("Enter Item Number.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    quantityTextBox.Focus();
                }
                //if it is not blank run this block
                else
                {
                    //try block to check if the format of what is input into the textbox is correct
                    try
                    {
                        //convert the string to an integer
                        quantity = Convert.ToInt32(quantityTextBox.Text);

                        //if the quantity is greater than 10 make groupbox1 visible
                        if (quantity >= 10)
                        {
                            groupBox1.Visible = true;
                        }

                        //calls the checkMultiplier method
                        checkMultiplier();

                        //disables the quantityTextbox, and puts the focus on the upcTextBox
                        quantityTextBox.Enabled = false;
                        upcTextBox.Focus();
                    }
                    //if user inputs the wrong format run this block
                    catch (FormatException)
                    {
                        //MessageBox to display an error message, clear, and place the cursor in the quantityTextBox
                        MessageBox.Show("Please Insert A Valid Quantity", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        quantityTextBox.Clear();
                        quantityTextBox.Focus();
                    }

                }
            }
        }

        //-----------END OF KeyDown events-----------//


        //-----------START OF checkMultiplier METHOD-----------//

        //method for calculating the quantity remaining and testing conditions for multiplier
        private void checkMultiplier()
        {
           
            //calculating the number of items left
            quantityRemaining = quantity - countResult;

            //if there are less than 100 items left
            if (quantityRemaining < 100) 
            {
                multiplierComboBox.Items.Remove("100");
            }

            //if there are less than 50 items left disable the selection
            if (quantityRemaining < 50)
            {
                multiplierComboBox.Items.Remove("50");   
            }

            //if there are less than 20 items left disable the selection
            if (quantityRemaining < 20)
            {
                multiplierComboBox.Items.Remove("20");
            }

            //if there are less than 10 items left disable the selection
            //Force the selection to the the index of 0
            if (quantityRemaining < 10)
            {
                multiplierComboBox.Items.Remove("10");
                multiplierComboBox.SelectedIndex = 0;
            }

            if (quantityRemaining == 0)
            {
                //sets the time that the operation was completed
                verificationEndTime = DateTime.Now.ToLongTimeString();

                //variable to hold the numerical end time used to calculate duration
                lVerfEnd = DateTime.Now;
                //calculating the duration
                TimeSpan time = lVerfEnd - lVerfStart;
                //set the duration into a string to be exported into excel 
                timeElapsed = new DateTime(time.Ticks).ToString(@"mm:ss");
                MessageBox.Show("Completed Correctly!");

                //calls the cleaningUp method
                cleaningUp();
            }
        }
        //-----------END OF checkMultiplier METHOD-----------//


        //-----------START OF writeToExcel METHOD-----------//

        //Method for writing into excel
        private void writeToExcel()
        {
            //sets the file name to be saved
            ExcelFilePath = user + "_" + DateTime.Now.ToString("yyyyMMdd_hhss") + ".xlsx";

            try
            {
                if (Tbl == null || Tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                Excel.Application excelApp = new Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Excel._Worksheet workSheet = excelApp.ActiveSheet;

                // column headings
                for (int i = 0; i < Tbl.Columns.Count; i++)
                {
                    workSheet.Cells[1, (i + 1)] = Tbl.Columns[i].ColumnName;

                    //formatting columns A and C to numbers to display correctly
                    workSheet.Cells[1, 1].EntireColumn.NumberFormat = "@";
                    workSheet.Cells[1, 3].EntireColumn.NumberFormat = "@";

                }

                // creating and populating the rows
                for (int i = 0; i < Tbl.Rows.Count; i++)
                {
                    for (int j = 0; j < Tbl.Columns.Count; j++)
                    {
                        workSheet.Cells[(i + 2), (j + 1)] = Tbl.Rows[i][j];
                    }
                }

                //formatting the cells to automatically widen            
                workSheet.Cells[1, 1].EntireColumn.Autofit();
                workSheet.Cells[1, 2].EntireColumn.Autofit();
                workSheet.Cells[1, 3].EntireColumn.Autofit();
                workSheet.Cells[1, 4].EntireColumn.Autofit();
                workSheet.Cells[1, 5].EntireColumn.Autofit();
                workSheet.Cells[1, 6].EntireColumn.Autofit();
                workSheet.Cells[1, 7].EntireColumn.Autofit();
                workSheet.Cells[1, 8].EntireColumn.Autofit();
                workSheet.Cells[1, 9].EntireColumn.Autofit();

                // check filepath, if ExcelFilePath isn't null and blank run this block
                if (ExcelFilePath != null && ExcelFilePath != "")
                {
                    try
                    {
                        //save the excel worksheet as the the ExcelFilePath as declared earlier
                        workSheet.SaveAs(ExcelFilePath);
                        //Quits Excel
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }

                else    // no filepath is given
                {
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }

        //-----------END OF writeToExcel METHOD-----------//

        
        //-----------START OF ToolStripMenuItem_Click EVENTS-----------//

        private void logOutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //hides this form, and creates a new instance of the Login form
            if (MessageBox.Show("Do you want to log out?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Hide();
                (new Login()).Show();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //closes the application
            if (MessageBox.Show("Do you want to exit?", "Exit", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
        //-----------END OF ToolStripMenuItem_Click EVENTS-----------//


        //if the main hub button is clicked run event
        private void mainhubButton_Click(object sender, EventArgs e)
        {

            //System.Diagnostics.Process.Start("chrome.exe", "gmail.com");

            //pop up a dialog box that asks if user wants to return to the main hub. 
            //if yes hide this form and create a new instance of the main hub form
            if (MessageBox.Show("Do you want to return to the Main Hub?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                this.Close();
            }
        }

        //-----------START OF cleaningUp METHOD-----------//

        private void cleaningUp()
        {
            //populating the row
            DataRow dr = Tbl.NewRow();
            dr["Delivery Number"] = deliveryNumber;
            dr["Item Number"] = item;
            dr["UPC"] = upc;
            dr["Total Quantity"] = quantity;
            dr["Quantity Scanned"] = countResult;
            dr["Quantity Remaining"] = quantityRemaining;
            dr["Start Time"] = verificationStartTime;
            dr["End Time"] = verificationEndTime;
            dr["Time Elapsed (mm:ss)"] = timeElapsed;

            //adding a row to the datatable
            Tbl.Rows.Add(dr);

            dataGridView1.DataSource = Tbl;

            //calls the resetValues method
            resetValues();
        }
        //-----------END OF cleanUp METHOD-----------//

    }
}
