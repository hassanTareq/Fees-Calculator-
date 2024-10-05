using MySql.Data.MySqlClient;
namespace WinFormsApp1Demo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            /*try
            {
                // establish database connection
                MySqlConnection con = new MySqlConnection();
                // connect to database
                con.ConnectionString = "server=localhost;uid=root;pwd=HassanTariq5000;database=register";
                con.Open();
                // send sql statement 
                MySqlCommand cmd = new MySqlCommand("SELECT Uname,Upassword FROM user", con);
                // define data reader 
                MySqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    if (this.textUserName.Text == "" + reader["Uname"] && this.textpassword.Text == "" + reader["Upassword"])
                    {
                        Register r = new Register();
                        r.Show();
                        this.Visible = false;

                    }
                    else
                    {
                        this.textUserName.Text = " ";
                        this.textpassword.Text = " ";
                        this.Refresh();
                        this.label1.Text = "wrong user name or password!";
                        this.label1.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }*/
            if (this.textUserName.Text == "hassan" && this.textpassword.Text == "Ais1998") 
            {
                Register r = new Register();
                r.Show();
                this.Visible = false;
            }
            else
            {
                this.textUserName.Text = " ";
                this.textpassword.Text = " ";
                this.Refresh();
                this.label1.Text = "wrong user name or password!";
                this.label1.Visible = true;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            this.button1.BackColor = Color.FromArgb(145, 40, 33);
            this.UserName.BackColor = Color.FromArgb(211, 185, 145);
            this.label2.BackColor = Color.FromArgb(211, 185, 145);
            this.password.BackColor = Color.FromArgb(211, 185, 145);
            this.label1.BackColor = Color.FromArgb(211, 185, 145);
            this.button1.ForeColor = Color.FromArgb(211, 185, 145);
        }
    }
}
