using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using Mysqlx;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Google.Protobuf.WellKnownTypes;

namespace WinFormsApp1Demo
{
    public partial class Register : Form
    {
        invoice rec;
        double registrationFee, testFee, activities, academicService, laptop, gradeFee, bookFee, busFee;
        double[] GradeFee, BookFee, BusFee;
        int cells;
        string excelfilePath, SheetName;
        // establish database connection
        /* MySqlConnection con = new MySqlConnection();
         MySqlCommand cmd;
         MySqlDataReader reader;*/
        Application app;
        Workbook existbook;
        public Register()
        {
            InitializeComponent();
            this.registrationFee = 1250.00;
            this.testFee = 1000.00;
            this.activities = 1000.00;
            this.academicService = 1350.00;
            this.laptop = 1600.00;
            cells = 15;
            this.GradeFee = [12500.00, 13500.00, 13750.00, 14000.00, 14500.00, 15000.00, 15500.00, 16000.00, 17000.00, 18000.00, 18500.00, 19500.00, 22000.00,23500.00];
            BookFee = [1150.00, 1250.0, 1500.00, 2400.00, 2400.00, 2400.00, 2400.00, 2400.00, 2500.00, 2650.00, 2950.00, 3500.00, 3500.00,3500.00];
            BusFee = [2800.00, 3500.00, 4500.00, 4500.00, 5500.00, 6500.00];
            this.SheetName = "Sheet2";
            /*
            // connect to database
            con.ConnectionString = "server=localhost;uid=root;pwd=HassanTariq5000;database=register";
            con.Open();*/
            // excelfilePath = "C:\\Users\\Admin\\Desktop\\اتفاقيه نافذ.xlsx";// clint
            excelfilePath = "C:\\Users\\hassa\\OneDrive\\Desktop\\اتفاقيه نافذ.xlsx";
            // excelfilePath = "C:\\Users\\User-02\\Desktop\\اتفاقيه نافذ.xlsx";//clint school
           // excelfilePath = "C:\\Users\\User015\\Desktop\\اتفاقيه نافذ.xlsx";//clint school 2
           // excelfilePath = "..\\Desktop\\اتفاقيه نافذ.xlsx";//clint school 2



        }
        private void resetVars() 
        {
            count = 1;
            this.comboBox1.SelectedIndex = -1;
            this.comboBox2.SelectedIndex = -1;// grade
            this.comboBox3.SelectedIndex = -1;// student status
            this.locationZons.SelectedIndex = -1;// locationzone
            this.numSiblings.Text = string.Empty;
            this.textBoxOnceDisc.Text = string.Empty;
            this.textDescount.Text = string.Empty;
            this.radioButtonNo.Checked = true;
            this.radioButton2.Checked = true;
            this.textFName.Text = string.Empty;//fname
            this.texLName.Text = string.Empty;//lname
            this.radioButtonLaptopYes.Checked = false;// checked
            this.radioButtonLaptopNo.Checked = false;// checked
            this.numSiblings.Text = string.Empty;// number of siblings
            this.radioButton1.Checked = false;// transportation non checked
            this.radioButton2.Checked = false;// transportation non checked
            discount = 5;
            cells = 15;
        }

        public class AutoWidthPanel : Panel
        {
            protected override void OnResize(EventArgs eventargs)
            {
                base.OnResize(eventargs);

                // Adjust the width of the panel to match the width of the parent form
                if (Parent != null)
                {
                    Width = Parent.ClientSize.Width;

                }
            }
        }
        int siblings = 0;
        int discount = 5, doubleIncrement = 2, count = 1;
        double regFee = 0.00;


        private void generate_Click(object sender, EventArgs e)
        {   // check if new student 
            if (this.comboBox3.SelectedIndex == 1)
            {   // check the grade from KG1 to KG3
                if (this.comboBox2.SelectedIndex == 0 || this.comboBox2.SelectedIndex == 1 || this.comboBox2.SelectedIndex == 2)
                {
                    // select the grade fees            SQL

                    /* cmd = new MySqlCommand("SELECT fees FROM grade WHERE grade= @val1", con);
                     cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                     //cmd.Prepare();
                     reader = cmd.ExecuteReader();
                     reader.Read();
                     this.gradeFee = double.Parse("" + reader["fees"]);
                     reader.Close();*/

                    this.gradeFee = GradeFee[this.comboBox2.SelectedIndex];
                    // select book fees

                    /* cmd = new MySqlCommand("SELECT fees FROM book WHERE gradeNmae= @val1", con);
                     cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                     //cmd.Prepare();
                     reader = cmd.ExecuteReader();
                     reader.Read();
                     this.bookFee = double.Parse("" + reader["fees"]);
                     reader.Close();*/
                    this.bookFee = BookFee[this.comboBox2.SelectedIndex];

                    // select bus fees

                    if (this.radioButton1.Checked)
                    {
                        /*cmd = new MySqlCommand("SELECT Busfees FROM Bus WHERE Location= @val1", con);
                        cmd.Parameters.AddWithValue("@val1", this.locationZons.SelectedItem);
                        //cmd.Prepare();
                        reader = cmd.ExecuteReader();
                        reader.Read();
                        this.busFee = double.Parse("" + reader["Busfees"]);
                        reader.Close();*/
                        this.busFee = BusFee[this.locationZons.SelectedIndex];
                    }
                    else
                        this.busFee = 0;
                    //--------------------------------------------------------------------------------end SQL 
                    //Saudi
                    if (this.comboBox1.SelectedIndex == 0)
                    {
                        if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {

                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                             0.00,// siblings discount
                                             0.00,// discount
                                             Int32.Parse(this.textBoxOnceDisc.Text),
                                             totalFeeFirst,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear table cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        // rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(15500.00,
                                             discount,// siblings discount
                                             0.00,// discount
                                             Int32.Parse(this.textBoxOnceDisc.Text),
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);

                                        // check the last student
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                            discount,// siblings discount
                                            0.00,// discount
                                            Int32.Parse(this.textBoxOnceDisc.Text),
                                            totalFee,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        // check the last student 
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }


                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFeeFirst,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        // clear table cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                           discount,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        // check last student
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                            discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                            Int32.Parse(this.textBoxOnceDisc.Text),
                                            totalFee,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = discount + (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        // check last student
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }
                            }// has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                // clear table cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /*rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           0.00,// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                // clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-------------------------------------------------------------------
                                resetVars();
                            }
                        }
                        else //without pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                            0.00,// siblings discount
                                            0.00,// discount
                                            0.00,
                                            totalFeeFirst,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = 0;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //  rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             0.00,// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             0.00,// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /*  rec = new invoice(this.gradeFee,
                                             0.00,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFeeFirst,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*   rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           Int32.Parse(this.textDescount.Text),// discount
                                           0.00,
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = this.gradeFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                          0.00,// siblings discount
                                          0.00,// discount
                                          0.00,
                                          totalFee,
                                          this.textFName.Text + " " + this.texLName.Text,// student name
                                          (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                          (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                          (this.activities + (this.activities * 15.00) / 100),// activities fees
                                          (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                          (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                          (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                          0.00,// vat
                                          this.bookFee// book fee
                                          );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = 0;
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-------------------------------------------------------------
                                resetVars();
                            }
                        }
                    }
                    //--------------------------------------------------------------------------------------------------------------------
                    // another nationality
                    else
                    {

                        if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           0.00,// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFeeFirst,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.testFee + (this.testFee * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           15,// vat
                                           (this.bookFee + (this.bookFee * 15) / 100)// bookfee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        // rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                          discount,// siblings discount
                                          0.00,// discount
                                          Int32.Parse(this.textBoxOnceDisc.Text),
                                          totalFee,
                                          this.textFName.Text + " " + this.texLName.Text,// student name
                                          (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                          (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                          (this.activities + (this.activities * 15.00) / 100),// activities fees
                                          (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                          (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                          (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                          15,// vat
                                          (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                          );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         0.00,// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         15,// vat
                                         (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                         );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /*  rec = new invoice(this.gradeFee,
                                         0.00,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         15,// vat
                                         (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                         );
                                  rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                //  open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-----------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text)) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //--------------------------------------------------------------
                                resetVars();
                            }

                        }
                        else //without pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = 0;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();
                                        //create 'workbook' object

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                        //create 'workbook' object

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                        //create 'workbook' object

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                       discount,// siblings discount
                                       Int32.Parse(this.textDescount.Text),// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                       );*/
                                        //create 'workbook' object
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        //create 'workbook' object
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = 0;
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-----------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                       0.00,// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                       );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = 0;
                                ws.Cell("A" + cells).Value = 0;
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-----------------------------------------------------------
                                resetVars();
                            }
                        }
                    }

                }
                // check if grade from G1 to G3 and new student
                else if (this.comboBox2.SelectedIndex == 3 || this.comboBox2.SelectedIndex == 4 || this.comboBox2.SelectedIndex == 5)
                {
                    // select the grade fees            SQL

                    /* cmd = new MySqlCommand("SELECT fees FROM grade WHERE grade= @val1", con);
                     cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                     //cmd.Prepare();
                     reader = cmd.ExecuteReader();
                     reader.Read();
                     this.gradeFee = double.Parse("" + reader["fees"]);
                     reader.Close();*/

                    this.gradeFee = GradeFee[this.comboBox2.SelectedIndex];
                    // select book fees

                    /* cmd = new MySqlCommand("SELECT fees FROM book WHERE gradeNmae= @val1", con);
                     cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                     //cmd.Prepare();
                     reader = cmd.ExecuteReader();
                     reader.Read();
                     this.bookFee = double.Parse("" + reader["fees"]);
                     reader.Close();*/
                    this.bookFee = BookFee[this.comboBox2.SelectedIndex];

                    // select bus fees

                    if (this.radioButton1.Checked)
                    {
                        /*cmd = new MySqlCommand("SELECT Busfees FROM Bus WHERE Location= @val1", con);
                        cmd.Parameters.AddWithValue("@val1", this.locationZons.SelectedItem);
                        //cmd.Prepare();
                        reader = cmd.ExecuteReader();
                        reader.Read();
                        this.busFee = double.Parse("" + reader["Busfees"]);
                        reader.Close();*/
                        this.busFee = BusFee[this.locationZons.SelectedIndex];
                    }
                    else
                        this.busFee = 0;
                    //--------------------------------------------------------------------------------end SQL 
                    //Saudi
                    if (this.comboBox1.SelectedIndex == 0)
                    {
                        if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {

                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                             0.00,// siblings discount
                                             0.00,// discount
                                             Int32.Parse(this.textBoxOnceDisc.Text),
                                             totalFeeFirst,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear table cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        // rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(15500.00,
                                             discount,// siblings discount
                                             0.00,// discount
                                             Int32.Parse(this.textBoxOnceDisc.Text),
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                            discount,// siblings discount
                                            0.00,// discount
                                            Int32.Parse(this.textBoxOnceDisc.Text),
                                            totalFee,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFeeFirst,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        // clear table cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                           discount,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                            discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                            Int32.Parse(this.textBoxOnceDisc.Text),
                                            totalFee,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount + (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }
                            }// has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                // clear table cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /*rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           0.00,// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                // clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //------------------------------------------------------------
                                resetVars();
                            }
                        }
                        else //without pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                            0.00,// siblings discount
                                            0.00,// discount
                                            0.00,
                                            totalFeeFirst,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = 0;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //  rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             0.00,// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             0.00,// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /*  rec = new invoice(this.gradeFee,
                                             0.00,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFeeFirst,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*   rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           Int32.Parse(this.textDescount.Text),// discount
                                           0.00,
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = this.gradeFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                          0.00,// siblings discount
                                          0.00,// discount
                                          0.00,
                                          totalFee,
                                          this.textFName.Text + " " + this.texLName.Text,// student name
                                          (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                          (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                          (this.activities + (this.activities * 15.00) / 100),// activities fees
                                          (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                          (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                          (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                          0.00,// vat
                                          this.bookFee// book fee
                                          );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = 0;
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //---------------------------------------------------------------------
                                resetVars();
                            }
                        }
                    }
                    //--------------------------------------------------------------------------------------------------------------------
                    // another nationality
                    else
                    {

                        if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           0.00,// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFeeFirst,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.testFee + (this.testFee * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           15,// vat
                                           (this.bookFee + (this.bookFee * 15) / 100)// bookfee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        // rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                          discount,// siblings discount
                                          0.00,// discount
                                          Int32.Parse(this.textBoxOnceDisc.Text),
                                          totalFee,
                                          this.textFName.Text + " " + this.texLName.Text,// student name
                                          (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                          (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                          (this.activities + (this.activities * 15.00) / 100),// activities fees
                                          (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                          (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                          (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                          15,// vat
                                          (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                          );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         0.00,// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         15,// vat
                                         (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                         );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /*  rec = new invoice(this.gradeFee,
                                         0.00,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         15,// vat
                                         (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                         );
                                  rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text)) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-----------------------------------------------------------
                                resetVars();
                            }

                        }
                        else //without pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = 0;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                       discount,// siblings discount
                                       Int32.Parse(this.textDescount.Text),// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                       );*/
                                        //create 'workbook' object
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        //create 'workbook' object
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        ws.Cell("S" + cells).Value = "0";
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                       0.00,// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                       );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = 0;
                                ws.Cell("S" + cells).Value = "0";
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-----------------------------------------------------------
                                resetVars(); 
                            }
                        }
                    }


                }
                else // grade from G4 to G10 and new student
                {
                    // select the grade fees            SQL

                    /* cmd = new MySqlCommand("SELECT fees FROM grade WHERE grade= @val1", con);
                     cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                     //cmd.Prepare();
                     reader = cmd.ExecuteReader();
                     reader.Read();
                     this.gradeFee = double.Parse("" + reader["fees"]);
                     reader.Close();*/

                    this.gradeFee = GradeFee[this.comboBox2.SelectedIndex];
                    // select book fees

                    /* cmd = new MySqlCommand("SELECT fees FROM book WHERE gradeNmae= @val1", con);
                     cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                     //cmd.Prepare();
                     reader = cmd.ExecuteReader();
                     reader.Read();
                     this.bookFee = double.Parse("" + reader["fees"]);
                     reader.Close();*/
                    this.bookFee = BookFee[this.comboBox2.SelectedIndex];

                    // select bus fees

                    if (this.radioButton1.Checked)
                    {
                        /*cmd = new MySqlCommand("SELECT Busfees FROM Bus WHERE Location= @val1", con);
                        cmd.Parameters.AddWithValue("@val1", this.locationZons.SelectedItem);
                        //cmd.Prepare();
                        reader = cmd.ExecuteReader();
                        reader.Read();
                        this.busFee = double.Parse("" + reader["Busfees"]);
                        reader.Close();*/
                        this.busFee = BusFee[this.locationZons.SelectedIndex];
                    }
                    else
                        this.busFee = 0;
                    //--------------------------------------------------------------------------------end SQL 
                    //Saudi
                    if (this.comboBox1.SelectedIndex == 0)
                    {
                        if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {

                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100)  + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                             0.00,// siblings discount
                                             0.00,// discount
                                             Int32.Parse(this.textBoxOnceDisc.Text),
                                             totalFeeFirst,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear table cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        // rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(15500.00,
                                             discount,// siblings discount
                                             0.00,// discount
                                             Int32.Parse(this.textBoxOnceDisc.Text),
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                            discount,// siblings discount
                                            0.00,// discount
                                            Int32.Parse(this.textBoxOnceDisc.Text),
                                            totalFee,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFeeFirst,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        // clear table cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                           discount,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                            discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                            Int32.Parse(this.textBoxOnceDisc.Text),
                                            totalFee,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount + (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }
                            }// has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                            Int32.Parse(this.textDescount.Text),// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                // clear table cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /*rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           0.00,// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                // clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------
                                resetVars();
                            }
                        }
                        else //without pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                            0.00,// siblings discount
                                            0.00,// discount
                                            0.00,
                                            totalFeeFirst,
                                            this.textFName.Text + " " + this.texLName.Text,// student name
                                            (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                            (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                            (this.activities + (this.activities * 15.00) / 100),// activities fees
                                            (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                            (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                            (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                            0.00,// vat
                                            this.bookFee// book fee
                                            );*/
                                        var wb = new XLWorkbook(excelfilePath);
                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cell
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = 0;
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //  rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             0.00,// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             0.00,// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /*  rec = new invoice(this.gradeFee,
                                             0.00,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFeeFirst,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*   rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        double totalFee = regFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                             discount,// siblings discount
                                             Int32.Parse(this.textDescount.Text),// discount
                                             0.00,
                                             totalFee,
                                             this.textFName.Text + " " + this.texLName.Text,// student name
                                             (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                             (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                             (this.activities + (this.activities * 15.00) / 100),// activities fees
                                             (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                             (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                             (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                             0.00,// vat
                                             this.bookFee// book fee
                                             );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        ws.Cell("G" + cells).Value = this.bookFee;
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           Int32.Parse(this.textDescount.Text),// discount
                                           0.00,
                                           totalFee,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           0.00,// vat
                                           this.bookFee// book fee
                                           );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = this.gradeFee + this.bookFee + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                          0.00,// siblings discount
                                          0.00,// discount
                                          0.00,
                                          totalFee,
                                          this.textFName.Text + " " + this.texLName.Text,// student name
                                          (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                          (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                          (this.activities + (this.activities * 15.00) / 100),// activities fees
                                          (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                          (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                          (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                          0.00,// vat
                                          this.bookFee// book fee
                                          );
                                 rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = this.gradeFee;
                                ws.Cell("G" + cells).Value = this.bookFee;
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = 0;
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------
                                resetVars();
                            }
                        }
                    }
                    //--------------------------------------------------------------------------------------------------------------------
                    // another nationality
                    else
                    {

                        if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                           0.00,// siblings discount
                                           0.00,// discount
                                           Int32.Parse(this.textBoxOnceDisc.Text),
                                           totalFeeFirst,
                                           this.textFName.Text + " " + this.texLName.Text,// student name
                                           (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                           (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                           (this.activities + (this.activities * 15.00) / 100),// activities fees
                                           (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                           (this.testFee + (this.testFee * 15.00) / 100),//laptop fees
                                           (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                           15,// vat
                                           (this.bookFee + (this.bookFee * 15) / 100)// bookfee
                                           );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        // rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                          discount,// siblings discount
                                          0.00,// discount
                                          Int32.Parse(this.textBoxOnceDisc.Text),
                                          totalFee,
                                          this.textFName.Text + " " + this.texLName.Text,// student name
                                          (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                          (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                          (this.activities + (this.activities * 15.00) / 100),// activities fees
                                          (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                          (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                          (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                          15,// vat
                                          (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                          );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        // rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         0.00,// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         15,// vat
                                         (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                         );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100)+ (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            cells++;
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet1";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /*  rec = new invoice(this.gradeFee,
                                         0.00,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         15,// vat
                                         (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                         );
                                  rec.Show();*/
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-------------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text)) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet1";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //-------------------------------------------------------------
                                resetVars();
                            }

                        }
                        else //without pay once discount
                        {
                            //new condition has siblings and didn't add specific discount
                            if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = 0;
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();
                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = discount;
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();
                                    }
                                }

                            }
                            //new condition has siblings and specific discount
                            else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                            {
                                if (siblings != 0)
                                {
                                    if (count == 1)
                                    {
                                        double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                        totalFeeFirst += totalFeeFirst * 15 / 100;
                                        var temp = totalFeeFirst;
                                        totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                        /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                        );*/
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);
                                        //clear cells
                                        ws.Range("A15:T19").Value = "";
                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFeeFirst += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFeeFirst;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                        //rec.Show();

                                    }
                                    else if (count == 2)
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /* rec = new invoice(this.gradeFee,
                                       discount,// siblings discount
                                       Int32.Parse(this.textDescount.Text),// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                       );*/
                                        //create 'workbook' object
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            discount *= 2;
                                            siblings--;
                                            cells++;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                    else
                                    {
                                        regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                        var temp = regFee + (regFee * 15) / 100;
                                        double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                        //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                        /*  rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );*/
                                        //create 'workbook' object
                                        var wb = new XLWorkbook(excelfilePath);

                                        SheetName = "Sheet2";
                                        //create 'worksheet' object
                                        var ws = wb.Worksheet(SheetName);

                                        //write cells
                                        ws.Cell("B" + cells).Value = this.textFName.Text;
                                        ws.Cell("D12").Value = this.texLName.Text;
                                        ws.Cell("E" + cells).Value = temp;
                                        ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                        ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                        ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                        ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                        ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                        ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                        ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                        if (this.radioButtonLaptopNo.Checked)
                                        {
                                            totalFee += 0;
                                            ws.Cell("S" + cells).Value = "0";
                                        }
                                        else if (this.radioButtonLaptopYes.Checked)
                                        {
                                            totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                            ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                        }
                                        ws.Cell("T" + cells).Value = totalFee;
                                        ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                        wb.SaveAs(excelfilePath);
                                        if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                        {
                                            this.app = new Application();
                                            app.Visible = true;
                                            existbook = app.Workbooks.Open(excelfilePath);
                                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                            resetVars();
                                        }
                                        else
                                        {
                                            MessageBox.Show($"Student {count} saved successfully");
                                            this.comboBox2.SelectedIndex = -1;// grade
                                            this.comboBox3.SelectedIndex = -1;// student status
                                            this.locationZons.SelectedIndex = -1;// locationzone
                                            this.textFName.Text = string.Empty;//fname
                                            this.radioButtonLaptopYes.Checked = false;// checked
                                            this.radioButtonLaptopNo.Checked = false;// checked
                                            siblings--;
                                            this.numSiblings.Text = siblings.ToString();
                                            count++;
                                        }
                                        //rec.Show();

                                    }
                                }
                            }
                            // has discount but no siblings
                            else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                            {
                                double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100);
                                totalFee += totalFee * 15 / 100;
                                var temp = totalFee;
                                totalFee += (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                //MessageBox.Show("total fees: " + totalFee);
                                /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                        0.00,
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        15,// vat
                                        (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                        );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = temp;
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //----------------------------------------------------------
                                resetVars();
                            }
                            //new condition has no siblings and discount
                            else
                            {
                                double totalFee = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + (this.registrationFee + (this.registrationFee * 15.00) / 100) + (this.testFee + (this.testFee * 15.00) / 100) + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + (this.busFee + (this.busFee * 15.00) / 100);
                                /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                       0.00,// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                       );
                                 rec.Show();*/
                                //create 'workbook' object
                                var wb = new XLWorkbook(excelfilePath);

                                SheetName = "Sheet2";
                                //create 'worksheet' object
                                var ws = wb.Worksheet(SheetName);
                                //clear cells
                                ws.Range("A15:T19").Value = "";
                                //write cells
                                ws.Cell("B" + cells).Value = this.textFName.Text;
                                ws.Cell("D12").Value = this.texLName.Text;
                                ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                ws.Cell("J" + cells).Value = (this.registrationFee + (this.registrationFee * 15.00) / 100);
                                ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                ws.Cell("L" + cells).Value = (this.testFee + (this.testFee * 15.00) / 100);
                                ws.Cell("A" + cells).Value = 0;
                                if (this.radioButtonLaptopNo.Checked)
                                {
                                    totalFee += 0;
                                    ws.Cell("S" + cells).Value = "0";
                                }
                                else if (this.radioButtonLaptopYes.Checked)
                                {
                                    totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                    ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                }
                                ws.Cell("T" + cells).Value = totalFee;
                                ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                wb.SaveAs(excelfilePath);
                                // open excel file 
                                this.app = new Application();
                                app.Visible = true;
                                existbook = app.Workbooks.Open(excelfilePath);
                                bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                //--------------------------------------------------------------
                                resetVars();
                            }
                        }
                    }
                }
            }
            else // current student 
            {
                // select the grade fees            SQL

                /* cmd = new MySqlCommand("SELECT fees FROM grade WHERE grade= @val1", con);
                 cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                 //cmd.Prepare();
                 reader = cmd.ExecuteReader();
                 reader.Read();
                 this.gradeFee = double.Parse("" + reader["fees"]);
                 reader.Close();*/

                this.gradeFee = GradeFee[this.comboBox2.SelectedIndex];
                // select book fees

                /* cmd = new MySqlCommand("SELECT fees FROM book WHERE gradeNmae= @val1", con);
                 cmd.Parameters.AddWithValue("@val1", this.comboBox2.SelectedItem);
                 //cmd.Prepare();
                 reader = cmd.ExecuteReader();
                 reader.Read();
                 this.bookFee = double.Parse("" + reader["fees"]);
                 reader.Close();*/
                this.bookFee = BookFee[this.comboBox2.SelectedIndex];

                // select bus fees

                if (this.radioButton1.Checked)
                {
                    /*cmd = new MySqlCommand("SELECT Busfees FROM Bus WHERE Location= @val1", con);
                    cmd.Parameters.AddWithValue("@val1", this.locationZons.SelectedItem);
                    //cmd.Prepare();
                    reader = cmd.ExecuteReader();
                    reader.Read();
                    this.busFee = double.Parse("" + reader["Busfees"]);
                    reader.Close();*/
                    this.busFee = BusFee[this.locationZons.SelectedIndex];
                }
                else
                    this.busFee = 0;
                //--------------------------------------------------------------------------------end SQL 
                //Saudi
                if (this.comboBox1.SelectedIndex == 0)
                {
                    if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                    {
                        //new condition has siblings and didn't add specific discount
                        if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                        {

                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                         0.00,// siblings discount
                                         0.00,// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFeeFirst,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear table cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    // rec.Show();
                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(15500.00,
                                         discount,// siblings discount
                                         0.00,// discount
                                         Int32.Parse(this.textBoxOnceDisc.Text),
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        cells++;
                                        siblings--;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    // rec.Show();
                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                        0.00,// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        0.00,// vat
                                        this.bookFee// book fee
                                        );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;
                                    
                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();
                                }
                            }

                        }
                        //new condition has siblings and specific discount
                        else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                       Int32.Parse(this.textBoxOnceDisc.Text),
                                       totalFeeFirst,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       0.00,// vat
                                       this.bookFee// book fee
                                       );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    // clear table cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    //rec.Show();
                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                       discount,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                       Int32.Parse(this.textBoxOnceDisc.Text),
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       0.00,// vat
                                       this.bookFee// book fee
                                       );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        cells++;
                                        siblings--;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    // rec.Show();
                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /*  rec = new invoice(this.gradeFee,
                                        discount,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                        Int32.Parse(this.textBoxOnceDisc.Text),
                                        totalFee,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        0.00,// vat
                                        this.bookFee// book fee
                                        );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = discount + (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();
                                }
                            }
                        }// has discount but no siblings
                        else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                        {
                            double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100) + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            //MessageBox.Show("total fees: " + totalFee);
                            /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                        Int32.Parse(this.textDescount.Text),// discount
                                       Int32.Parse(this.textBoxOnceDisc.Text),
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       0.00,// vat
                                       this.bookFee// book fee
                                       );
                             rec.Show();*/
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet1";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            // clear table cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                            ws.Cell("G" + cells).Value = this.bookFee;
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                        //new condition has no siblings and discount
                        else
                        {
                            double totalFee = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100) + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            //MessageBox.Show("total fees: " + totalFee);
                            /*rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                       0.00,// discount
                                       Int32.Parse(this.textBoxOnceDisc.Text),
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       0.00,// vat
                                       this.bookFee// book fee
                                       );
                            rec.Show();*/
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet1";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            // clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                            ws.Cell("G" + cells).Value = this.bookFee;
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                    }
                    else //without pay once discount
                    {
                        //new condition has siblings and didn't add specific discount
                        if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                        0.00,// siblings discount
                                        0.00,// discount
                                        0.00,
                                        totalFeeFirst,
                                        this.textFName.Text + " " + this.texLName.Text,// student name
                                        (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                        (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                        (this.activities + (this.activities * 15.00) / 100),// activities fees
                                        (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                        (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                        (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                        0.00,// vat
                                        this.bookFee// book fee
                                        );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = 0;
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    //  rec.Show();
                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         0.00,// discount
                                         0.00,
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = discount;
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        cells++;
                                        siblings--;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    // rec.Show();

                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /*  rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         0.00,// discount
                                         0.00,
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * discount) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = discount;
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                            }

                        }
                        //new condition has siblings and specific discount
                        else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100) + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /*  rec = new invoice(this.gradeFee,
                                         0.00,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                         0.00,
                                         totalFeeFirst,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    //rec.Show();

                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /*   rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                         0.00,
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                    double totalFee = regFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /*  rec = new invoice(this.gradeFee,
                                         discount,// siblings discount
                                         Int32.Parse(this.textDescount.Text),// discount
                                         0.00,
                                         totalFee,
                                         this.textFName.Text + " " + this.texLName.Text,// student name
                                         (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                         (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                         (this.activities + (this.activities * 15.00) / 100),// activities fees
                                         (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                         (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                         (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                         0.00,// vat
                                         this.bookFee// book fee
                                         );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                    ws.Cell("G" + cells).Value = this.bookFee;
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    // rec.Show();

                                }
                            }
                        }
                        // has discount but no siblings
                        else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                        {
                            double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100) + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                       Int32.Parse(this.textDescount.Text),// discount
                                       0.00,
                                       totalFee,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       0.00,// vat
                                       this.bookFee// book fee
                                       );
                             rec.Show();*/
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet2";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            //clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100;
                            ws.Cell("G" + cells).Value = this.bookFee;
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                        //new condition has no siblings and discount
                        else
                        {
                            double totalFee = this.gradeFee + this.bookFee + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            //MessageBox.Show("total fees: " + totalFee);
                            /* rec = new invoice(this.gradeFee,
                                      0.00,// siblings discount
                                      0.00,// discount
                                      0.00,
                                      totalFee,
                                      this.textFName.Text + " " + this.texLName.Text,// student name
                                      (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                      (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                      (this.activities + (this.activities * 15.00) / 100),// activities fees
                                      (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                      (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                      (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                      0.00,// vat
                                      this.bookFee// book fee
                                      );
                             rec.Show();*/
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet2";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            //clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = this.gradeFee;
                            ws.Cell("G" + cells).Value = this.bookFee;
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = 0;
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                    }
                }
                //--------------------------------------------------------------------------------------------------------------------
                // another nationality
                else
                {

                    if (this.textBoxOnceDisc.Text.Length != 0)//with pay once discount
                    {
                        //new condition has siblings and didn't add specific discount
                        if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text) / 100);
                                    totalFeeFirst += totalFeeFirst * 15 / 100;
                                    var temp = totalFeeFirst;
                                    totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                       0.00,// siblings discount
                                       0.00,// discount
                                       Int32.Parse(this.textBoxOnceDisc.Text),
                                       totalFeeFirst,
                                       this.textFName.Text + " " + this.texLName.Text,// student name
                                       (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                       (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                       (this.activities + (this.activities * 15.00) / 100),// activities fees
                                       (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                       (this.testFee + (this.testFee * 15.00) / 100),//laptop fees
                                       (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                       15,// vat
                                       (this.bookFee + (this.bookFee * 15) / 100)// bookfee
                                       );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    // rec.Show();

                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                      discount,// siblings discount
                                      0.00,// discount
                                      Int32.Parse(this.textBoxOnceDisc.Text),
                                      totalFee,
                                      this.textFName.Text + " " + this.texLName.Text,// student name
                                      (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                      (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                      (this.activities + (this.activities * 15.00) / 100),// activities fees
                                      (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                      (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                      (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                      15,// vat
                                      (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                      );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        cells++;
                                        siblings--;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    // rec.Show();

                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text))) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                     discount,// siblings discount
                                     0.00,// discount
                                     Int32.Parse(this.textBoxOnceDisc.Text),
                                     totalFee,
                                     this.textFName.Text + " " + this.texLName.Text,// student name
                                     (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                     (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                     (this.activities + (this.activities * 15.00) / 100),// activities fees
                                     (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                     (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                     (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                     15,// vat
                                     (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                     );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + discount);
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                            }

                        }
                        //new condition has siblings and specific discount
                        else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text)) / 100);
                                    totalFeeFirst += totalFeeFirst * 15 / 100;
                                    var temp = totalFeeFirst;
                                    totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                    0.00,// siblings discount
                                    Int32.Parse(this.textDescount.Text),// discount
                                    Int32.Parse(this.textBoxOnceDisc.Text),
                                    totalFeeFirst,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    //rec.Show();

                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                    discount,// siblings discount
                                    Int32.Parse(this.textDescount.Text),// discount
                                    Int32.Parse(this.textBoxOnceDisc.Text),
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        cells++;
                                        siblings--;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                    discount,// siblings discount
                                    Int32.Parse(this.textDescount.Text),// discount
                                    Int32.Parse(this.textBoxOnceDisc.Text),
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet1";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                            }
                        }
                        // has discount but no siblings
                        else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                        {
                            double totalFee = (this.gradeFee - (this.gradeFee * (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text))) / 100);
                            totalFee += totalFee * 15 / 100;
                            var temp = totalFee;
                            totalFee += (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            //MessageBox.Show("total fees: " + totalFee);
                            /*  rec = new invoice(this.gradeFee,
                                     0.00,// siblings discount
                                     Int32.Parse(this.textDescount.Text),// discount
                                     Int32.Parse(this.textBoxOnceDisc.Text),
                                     totalFee,
                                     this.textFName.Text + " " + this.texLName.Text,// student name
                                     (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                     (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                     (this.activities + (this.activities * 15.00) / 100),// activities fees
                                     (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                     (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                     (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                     15,// vat
                                     (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                     );
                              rec.Show();*/
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet1";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            //clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = temp;
                            ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = (Int32.Parse(this.textBoxOnceDisc.Text) + Int32.Parse(this.textDescount.Text));
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                        //new condition has no siblings and discount
                        else
                        {
                            double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textBoxOnceDisc.Text)) / 100);
                            totalFee += totalFee * 15 / 100;
                            var temp = totalFee;
                            totalFee += (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            //MessageBox.Show("total fees: " + totalFee);
                            /* rec = new invoice(this.gradeFee,
                                    0.00,// siblings discount
                                    0.00,// discount
                                    Int32.Parse(this.textBoxOnceDisc.Text),
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                    );
                             rec.Show();*/
                            //create 'workbook' object
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet1";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            //clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = temp;
                            ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = Int32.Parse(this.textBoxOnceDisc.Text);
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }

                    }
                    else //without pay once discount
                    {
                        //new condition has siblings and didn't add specific discount
                        if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length == 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                    0.00,// siblings discount
                                    0.00,// discount
                                    0.00,
                                    totalFeeFirst,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = 0;
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    //rec.Show();
                                    //create 'workbook' object

                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                    discount,// siblings discount
                                    0.00,// discount
                                    0.00,
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = discount;
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();
                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * discount) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                    discount,// siblings discount
                                    0.00,// discount
                                    0.00,
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = discount;
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();
                                }
                            }

                        }
                        //new condition has siblings and specific discount
                        else if (this.numSiblings.Text.Length != 0 && this.textDescount.Text.Length != 0)
                        {
                            if (siblings != 0)
                            {
                                if (count == 1)
                                {
                                    double totalFeeFirst = this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text) / 100);
                                    totalFeeFirst += totalFeeFirst * 15 / 100;
                                    var temp = totalFeeFirst;
                                    totalFeeFirst += (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    // MessageBox.Show("total fees: for student number 1: " + totalFeeFirst);
                                    /* rec = new invoice(this.gradeFee,
                                    0.00,// siblings discount
                                    Int32.Parse(this.textDescount.Text),// discount
                                    0.00,
                                    totalFeeFirst,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                    );*/
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);
                                    //clear cells
                                    ws.Range("A15:T19").Value = "";
                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFeeFirst += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFeeFirst += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFeeFirst;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    MessageBox.Show($"Student {count} saved successfully");
                                    this.comboBox2.SelectedIndex = -1;// grade
                                    this.comboBox3.SelectedIndex = -1;// student status
                                    this.locationZons.SelectedIndex = -1;// locationzone
                                    this.textFName.Text = string.Empty;//fname
                                    this.radioButtonLaptopYes.Checked = false;// checked
                                    this.radioButtonLaptopNo.Checked = false;// checked
                                    siblings--;
                                    cells++;
                                    this.numSiblings.Text = siblings.ToString();
                                    count++;
                                    //rec.Show();

                                }
                                else if (count == 2)
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /* rec = new invoice(this.gradeFee,
                                   discount,// siblings discount
                                   Int32.Parse(this.textDescount.Text),// discount
                                   0.00,
                                   totalFee,
                                   this.textFName.Text + " " + this.texLName.Text,// student name
                                   (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                   (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                   (this.activities + (this.activities * 15.00) / 100),// activities fees
                                   (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                   (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                   (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                   15,// vat
                                   (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                   );*/
                                    //create 'workbook' object
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        discount *= 2;
                                        siblings--;
                                        cells++;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                                else
                                {
                                    regFee = this.gradeFee - (this.gradeFee * (discount + Int32.Parse(this.textDescount.Text))) / 100;
                                    var temp = regFee + (regFee * 15) / 100;
                                    double totalFee = (regFee + (regFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                                    //MessageBox.Show($"total fees for student number {count}: {totalFee}");
                                    /*  rec = new invoice(this.gradeFee,
                                    discount,// siblings discount
                                    Int32.Parse(this.textDescount.Text),// discount
                                    0.00,
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                    );*/
                                    //create 'workbook' object
                                    var wb = new XLWorkbook(excelfilePath);

                                    SheetName = "Sheet2";
                                    //create 'worksheet' object
                                    var ws = wb.Worksheet(SheetName);

                                    //write cells
                                    ws.Cell("B" + cells).Value = this.textFName.Text;
                                    ws.Cell("D12").Value = this.texLName.Text;
                                    ws.Cell("E" + cells).Value = temp;
                                    ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                                    ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                                    ws.Cell("J" + cells).Value = 0;
                                    ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                                    ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                                    ws.Cell("L" + cells).Value = 0;
                                    ws.Cell("A" + cells).Value = (discount + Int32.Parse(this.textDescount.Text));
                                    if (this.radioButtonLaptopNo.Checked)
                                    {
                                        totalFee += 0;
                                        ws.Cell("S" + cells).Value = "0";
                                    }
                                    else if (this.radioButtonLaptopYes.Checked)
                                    {
                                        totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                        ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                                    }
                                    ws.Cell("T" + cells).Value = totalFee;
                                    ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                                    wb.SaveAs(excelfilePath);
                                    if (Int128.Parse(this.numSiblings.Text) - 1 == 0)
                                    {
                                        this.app = new Application();
                                        app.Visible = true;
                                        existbook = app.Workbooks.Open(excelfilePath);
                                        bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                                        resetVars();
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Student {count} saved successfully");
                                        this.comboBox2.SelectedIndex = -1;// grade
                                        this.comboBox3.SelectedIndex = -1;// student status
                                        this.locationZons.SelectedIndex = -1;// locationzone
                                        this.textFName.Text = string.Empty;//fname
                                        this.radioButtonLaptopYes.Checked = false;// checked
                                        this.radioButtonLaptopNo.Checked = false;// checked
                                        siblings--;
                                        this.numSiblings.Text = siblings.ToString();
                                        count++;
                                    }
                                    //rec.Show();

                                }
                            }
                        }
                        // has discount but no siblings
                        else if (this.numSiblings.Text.Length == 0 && this.textDescount.Text.Length != 0)
                        {
                            double totalFee = (this.gradeFee - (this.gradeFee * Int32.Parse(this.textDescount.Text)) / 100);
                            totalFee += totalFee * 15 / 100;
                            var temp = totalFee;
                            totalFee += (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            //MessageBox.Show("total fees: " + totalFee);
                            /* rec = new invoice(this.gradeFee,
                                    0.00,// siblings discount
                                    Int32.Parse(this.textDescount.Text),// discount
                                    0.00,
                                    totalFee,
                                    this.textFName.Text + " " + this.texLName.Text,// student name
                                    (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                    (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                    (this.activities + (this.activities * 15.00) / 100),// activities fees
                                    (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                    (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                    (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                    15,// vat
                                    (this.bookFee + (this.bookFee * 15) / 100)// book fee 
                                    );
                             rec.Show();*/
                            //create 'workbook' object
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet2";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            //clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = temp;
                            ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = Int32.Parse(this.textDescount.Text);
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                            //MessageBox.Show("Check اتفاقيه نافذ to print the invoice");

                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                        //new condition has no siblings and discount
                        else
                        {
                            double totalFee = (this.gradeFee + (this.gradeFee * 15) / 100) + (this.bookFee + (this.bookFee * 15) / 100) + 0 + 0 + (this.activities + (this.activities * 15.00) / 100) + (this.academicService + (this.academicService * 15.00) / 100) + 0 + (this.busFee + (this.busFee * 15.00) / 100);
                            /* rec = new invoice(this.gradeFee,
                                   0.00,// siblings discount
                                   0.00,// discount
                                   0.00,
                                   totalFee,
                                   this.textFName.Text + " " + this.texLName.Text,// student name
                                   (this.registrationFee + (this.registrationFee * 15.00) / 100), //reg fee
                                   (this.testFee + (this.testFee * 15.00) / 100),// test fees
                                   (this.activities + (this.activities * 15.00) / 100),// activities fees
                                   (this.academicService + (this.academicService * 15.00) / 100),// academic service fees
                                   (this.laptop + (this.laptop * 15.00) / 100),//laptop fees
                                   (this.busFee + (this.busFee * 15.00) / 100),//bus fees
                                   15,// vat
                                   (this.bookFee + (this.bookFee * 15) / 100)// book fee
                                   );
                             rec.Show();*/
                            //create 'workbook' object
                            var wb = new XLWorkbook(excelfilePath);

                            SheetName = "Sheet2";
                            //create 'worksheet' object
                            var ws = wb.Worksheet(SheetName);
                            //clear cells
                            ws.Range("A15:T19").Value = "";
                            //write cells
                            ws.Cell("B" + cells).Value = this.textFName.Text;
                            ws.Cell("D12").Value = this.texLName.Text;
                            ws.Cell("E" + cells).Value = (this.gradeFee + (this.gradeFee * 15) / 100);
                            ws.Cell("G" + cells).Value = (this.bookFee + (this.bookFee * 15) / 100);
                            ws.Cell("I" + cells).Value = (this.busFee + (this.busFee * 15.00) / 100);
                            ws.Cell("J" + cells).Value = 0;
                            ws.Cell("N" + cells).Value = (this.activities + (this.activities * 15.00) / 100);
                            ws.Cell("P" + cells).Value = (this.academicService + (this.academicService * 15.00) / 100);
                            ws.Cell("L" + cells).Value = 0;
                            ws.Cell("A" + cells).Value = 0;
                            if (this.radioButtonLaptopNo.Checked)
                            {
                                totalFee += 0;
                                ws.Cell("S" + cells).Value = "0";
                            }
                            else if (this.radioButtonLaptopYes.Checked)
                            {
                                totalFee += (this.laptop + (this.laptop * 15.00) / 100);
                                ws.Cell("S" + cells).Value = "" + (this.laptop + (this.laptop * 15.00) / 100);
                            }
                            ws.Cell("T" + cells).Value = totalFee;
                            ws.Cell("D" + cells).Value = "" + this.comboBox2.SelectedItem;

                            wb.SaveAs(excelfilePath);
                           // MessageBox.Show("Check اتفاقيه نافذ to print the invoice");
                            // auto open excel file 
                            this.app = new Application();
                            app.Visible = true;
                            existbook = app.Workbooks.Open(excelfilePath);
                            bool excelWorksheet = existbook.Sheets[SheetName].Activate;
                            //--------------------------------------------------------------
                            resetVars();
                        }
                    }
                }//--

            }



        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Form1 f = new Form1();
            f.Visible = true;
            f.Refresh();
            this.Close();

        }

        private void numSiblings_TextChanged(object sender, EventArgs e)
        {
            if (this.numSiblings.Text.Length != 0)
                siblings = Int32.Parse(this.numSiblings.Text);
        }



        private void radioButtonYes_CheckedChanged(object sender, EventArgs e)
        {
            this.label3.Visible = true;
            this.numSiblings.Visible = true;
            this.label6.Visible = true;
            this.textDescount.Visible = true;
        }

        private void radioButtonNo_CheckedChanged(object sender, EventArgs e)
        {
            this.label3.Visible = false;
            this.numSiblings.Visible = false;
            //this.label6.Visible = false;
            //this.textDescount.Visible = false;
            this.numSiblings.Text = string.Empty;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            this.label11.Visible = true;
            this.locationZons.Visible = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            this.label11.Visible = false;
            this.locationZons.Visible = false;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            count = 1;
            this.comboBox1.SelectedIndex = -1;
            this.comboBox2.SelectedIndex = -1;// grade
            this.comboBox3.SelectedIndex = -1;// student status
            this.locationZons.SelectedIndex = -1;// locationzone
            this.numSiblings.Text = string.Empty;
            this.textBoxOnceDisc.Text = string.Empty;
            this.textDescount.Text = string.Empty;
            this.radioButtonNo.Checked = true;
            this.radioButton2.Checked = true;
            this.textFName.Text = string.Empty;//fname
            this.texLName.Text = string.Empty;//lname
            this.radioButtonLaptopYes.Checked = false;// checked
            this.radioButtonLaptopNo.Checked = false;// checked
            discount = 5;
            cells = 15;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            this.btnLogout.BackColor = Color.FromArgb(145, 40, 33);
            this.btnNew.BackColor = Color.FromArgb(145, 40, 33);
            this.btnNew.ForeColor = Color.White;
            this.btnExcel.BackColor = Color.FromArgb(145, 40, 33);
            this.btnExcel.ForeColor = Color.White;
            this.generate.BackColor = Color.FromArgb(145, 40, 33);
            this.generate.ForeColor = Color.White;
            this.label1.BackColor = Color.FromArgb(211, 185, 145);
            this.label12.BackColor = Color.FromArgb(211, 185, 145);
            this.radioButtonYes.BackColor = Color.FromArgb(211, 185, 145);
            this.radioButtonNo.BackColor = Color.FromArgb(211, 185, 145);
            this.label3.BackColor = Color.FromArgb(211, 185, 145);
            this.label5.BackColor = Color.FromArgb(211, 185, 145);
            this.label6.BackColor = Color.FromArgb(211, 185, 145);
            this.label4.BackColor = Color.FromArgb(211, 185, 145);
            this.label13.BackColor = Color.FromArgb(211, 185, 145);
            this.label7.BackColor = Color.FromArgb(211, 185, 145);
            this.label8.BackColor = Color.FromArgb(211, 185, 145);
            this.label9.BackColor = Color.FromArgb(211, 185, 145);
            this.label10.BackColor = Color.FromArgb(211, 185, 145);
            this.label11.BackColor = Color.FromArgb(211, 185, 145);
            this.radioButtonLaptopYes.BackColor = Color.FromArgb(211, 185, 145);
            this.radioButtonLaptopNo.BackColor = Color.FromArgb(211, 185, 145);
            this.radioButton1.BackColor = Color.FromArgb(211, 185, 145);
            this.radioButton2.BackColor = Color.FromArgb(211, 185, 145);
            this.label2.BackColor = Color.FromArgb(211, 185, 145);
            this.label15.ForeColor = Color.FromArgb(211, 185, 145);
            this.label14.ForeColor = Color.FromArgb(211, 185, 145);
            this.label15.BackColor = Color.FromArgb(145, 40, 33);
            this.label14.BackColor = Color.FromArgb(145, 40, 33);
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            this.app = new Application();
            app.Visible = true;
            existbook = app.Workbooks.Open(excelfilePath);
            bool excelWorksheet = existbook.Sheets[SheetName].Activate;


        }
    }
}
