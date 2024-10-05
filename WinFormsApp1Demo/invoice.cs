using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.codec;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace WinFormsApp1Demo
{
    public partial class invoice : Form
    {
        double fees, sibDiscount, Discount, PayOnceDiscount, TotalFee, regfee, test, activities, academicService, laptop, bus, vat, book;
        string Stname,pdf="";

        public invoice(double fees, double sibDiscount, double Discount, double PayOnceDiscount, double TotalFee, string Stname, double regfee,
            double test, double activities, double academicService, double laptop, double bus, double vat, double book)
        {
            InitializeComponent();
            this.fees = fees;
            this.sibDiscount = sibDiscount;
            this.Discount = Discount;
            this.PayOnceDiscount = PayOnceDiscount;
            this.TotalFee = TotalFee;
            this.Stname = Stname;
            this.regfee = regfee;
            this.test = test;
            this.activities = activities;
            this.academicService = academicService;
            this.laptop = laptop;
            this.bus = bus;
            this.vat = vat;
            this.book = book;
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            this.name.Text = Stname;
            this.OriginalFees.Text = fees.ToString("F3") + " SR";
            this.SibDiscount.Text = sibDiscount.ToString() + " %";
            this.discount.Text = Discount.ToString() + " %";
            this.POD.Text = PayOnceDiscount.ToString() + " %";
            this.totalFee.Text = TotalFee.ToString() + " SR";
            this.Regfee.Text = regfee.ToString() + " SR";
            this.Testfee.Text = test.ToString() + " SR";
            this.Activitiesfee.Text = activities.ToString() + " SR";
            this.Acservice.Text = academicService.ToString() + " SR";
            this.Laptopfee.Text = laptop.ToString() + " SR";
            this.Busfee.Text = bus.ToString() + " SR";
            this.label4.Text = book.ToString() + " SR";
            pdf += $"\n Student name:       {Stname}\n Registration fee +15%      {regfee} SR\n";
            pdf += $"\n Fees:       {fees} SR\n Test fee +15%        {test}\n";
            pdf += $"\n Siblings discount:      {sibDiscount} SR\n Activities fee +15%        {activities} SR\n";
            pdf += $"\n New discount:       {Discount} SR\n Academic services +15%       {academicService} SR\n";
            pdf += $"\n Pay once discount:      {PayOnceDiscount} SR\n Laptop fee +15%       {laptop}  SR\n";
            if (vat != 0)
            {
                this.label13.Text = $"Book Fees + {vat.ToString()}%";
                this.label6.Text = $"Total Fees + {vat.ToString()}%";
                pdf += $"\n Book Fees +15%      {book} SR\n Bus fee +15%     {laptop} SR";
                pdf += $"\nTotal Fees + 15%     {TotalFee} SR";

            }
            else
            {
               pdf += $"\n Book Fees        {book} SR\n Bus fee +15%     {bus} SR";
               pdf += $"\n Total Fees       {TotalFee} SR";

            }
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            using(SaveFileDialog sfd= new SaveFileDialog() { Filter="PDF file|*.pdf", ValidateNames = true }) 
            {
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    //BaseFont baseFontBold = BaseFont.CreateFont(@"C:\Windows\Fonts\Calibrib.ttf", "Identity-H", BaseFont.EMBEDDED);
                    //var header = new iTextSharp.text.Font(baseFontBold, 12f);
                    //add image
                   iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance("C:\\Users\\hassa\\OneDrive\\Pictures\\school-removebg-preview.png");
                    image.ScaleToFit(140f,120f);
                    image.SpacingBefore = 10f;
                    image.SpacingAfter = 0f;
                    image.Alignment = Element.WRITABLE_DIRECT| Element.ALIGN_RIGHT;
                    //receipt details                           
                    var color1 = new BaseColor(93, 129, 234);
                    var price = new BaseColor(255, 0, 0);
                    var totalprice = new BaseColor(89, 142, 110);
                    var header = FontFactory.GetFont("Arial", 19,color1);
                    var school = FontFactory.GetFont("Arial", 15, price);
                    var important = FontFactory.GetFont("Arial", 15, price);
                    var body = FontFactory.GetFont("Arial", 15);
                    var gradeFont = FontFactory.GetFont("Arial", 15,color1);
                    var total = FontFactory.GetFont("Arial", 15, totalprice);
                    iTextSharp.text.Rectangle layout = new iTextSharp.text.Rectangle(PageSize.A4);
                    layout.BackgroundColor=new BaseColor(241, 234, 219);
                    iTextSharp.text.Document doc = new iTextSharp.text.Document(layout);
                    try
                    {
                        PdfWriter.GetInstance(doc, new FileStream(sfd.FileName, FileMode.Create));
                        doc.Open();
                        // add school name
                        Phrase schoolname = new Phrase();
                        //schoolname.Alignment = Element.ALIGN_RIGHT;
                        schoolname.Add(new iTextSharp.text.Phrase("\nAL AFAQ INTERNATIONAL SCHOOL",school));
                        schoolname.Add(new iTextSharp.text.Phrase("\nOwned & Supervised by"));
                        schoolname.Add(new iTextSharp.text.Phrase("\nAl-Feras Educational Company"));
                        schoolname.Add(new iTextSharp.text.Phrase("\nJeddah International School"));
                        schoolname.Add(new iTextSharp.text.Phrase("\nA.Y. 2024-2025"));
                        //-----------------------------------------------------
                        Paragraph p1 = new Paragraph("\n\nReceipt",header);
                        p1.Alignment = Element.ALIGN_TOP;
                        p1.Alignment= Element.ALIGN_CENTER;
                        Paragraph p2 = new Paragraph();
                        p2.Alignment = Element.ALIGN_CENTER;
                        p2.Add(new iTextSharp.text.Phrase("\n\n\n Student name:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {Stname}", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n School Grade:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    G4", gradeFont));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Registration fee +15%:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {regfee}", important));
                        p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Fees:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {fees}", important));
                        p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Test fee +15%:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {test}", important));
                        p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Siblings discount:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {sibDiscount}", important));
                        p2.Add(new iTextSharp.text.Phrase(" %", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Activities fee +15%:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {activities}", important));
                        p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n New discount:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {Discount}", important));
                        p2.Add(new iTextSharp.text.Phrase(" %", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Academic services +15%:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {academicService}", important));
                        p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Pay once discount:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {PayOnceDiscount}", important));
                        p2.Add(new iTextSharp.text.Phrase(" %", body));
                        p2.Add(new iTextSharp.text.Phrase("\n\n Laptop fee +15%t:", body));
                        p2.Add(new iTextSharp.text.Phrase($"    {laptop}", important));
                        p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        if (vat != 0)
                        {
                            p2.Add(new iTextSharp.text.Phrase("\n\n Book Fees +15%:", body));
                            p2.Add(new iTextSharp.text.Phrase($"    {book}", important));
                            p2.Add(new iTextSharp.text.Phrase(" SR", body));
                            p2.Add(new iTextSharp.text.Phrase("\n\n Total Fees + 15%:", body));
                            p2.Add(new iTextSharp.text.Phrase($"    {TotalFee}", total));
                            p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        }
                        else 
                        {
                            p2.Add(new iTextSharp.text.Phrase("\n\n Book Fees:", body));
                            p2.Add(new iTextSharp.text.Phrase($"    {book}", important));
                            p2.Add(new iTextSharp.text.Phrase(" SR", body));
                            p2.Add(new iTextSharp.text.Phrase("\n\n Total Fees:", body));
                            p2.Add(new iTextSharp.text.Phrase($"    {TotalFee}", total));
                            p2.Add(new iTextSharp.text.Phrase(" SR", body));
                        }
                        doc.Add(image);
                        doc.Add(schoolname);
                        doc.Add(p1);
                        doc.Add(p2);
                        //doc.Add(new iTextSharp.text.Paragraph(pdf));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally 
                    {
                        doc.Close();
                    }
                }
            } 
        }
    }
}
