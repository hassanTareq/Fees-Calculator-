namespace WinFormsApp1Demo
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            panel1 = new Panel();
            label2 = new Label();
            label1 = new Label();
            button1 = new Button();
            password = new Label();
            textpassword = new TextBox();
            UserName = new Label();
            textUserName = new TextBox();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.AutoSize = true;
            panel1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            panel1.BackgroundImage = Properties.Resources.login2;
            panel1.BackgroundImageLayout = ImageLayout.Stretch;
            panel1.Controls.Add(label2);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(button1);
            panel1.Controls.Add(password);
            panel1.Controls.Add(textpassword);
            panel1.Controls.Add(UserName);
            panel1.Controls.Add(textUserName);
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(800, 450);
            panel1.TabIndex = 0;
            panel1.Paint += panel1_Paint;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            label2.Location = new Point(405, 56);
            label2.Name = "label2";
            label2.Size = new Size(55, 23);
            label2.TabIndex = 7;
            label2.Text = "Login";
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            label1.AutoSize = true;
            label1.ForeColor = Color.Red;
            label1.Location = new Point(304, 238);
            label1.Name = "label1";
            label1.Size = new Size(50, 20);
            label1.TabIndex = 6;
            label1.Text = "label1";
            label1.Visible = false;
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom;
            button1.BackColor = Color.White;
            button1.Cursor = Cursors.Hand;
            button1.FlatStyle = FlatStyle.Popup;
            button1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            button1.ForeColor = SystemColors.ActiveCaptionText;
            button1.Location = new Point(388, 275);
            button1.Name = "button1";
            button1.Size = new Size(94, 29);
            button1.TabIndex = 4;
            button1.Text = "Login";
            button1.UseVisualStyleBackColor = false;
            button1.Click += button1_Click;
            button1.Enter += button1_Click;
            // 
            // password
            // 
            password.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            password.AutoSize = true;
            password.Font = new Font("Segoe UI", 10F);
            password.Location = new Point(302, 179);
            password.Name = "password";
            password.Size = new Size(81, 23);
            password.TabIndex = 3;
            password.Text = "password";
            // 
            // textpassword
            // 
            textpassword.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            textpassword.Location = new Point(405, 175);
            textpassword.Name = "textpassword";
            textpassword.Size = new Size(125, 27);
            textpassword.TabIndex = 2;
            textpassword.UseSystemPasswordChar = true;
            // 
            // UserName
            // 
            UserName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            UserName.AutoSize = true;
            UserName.BackColor = Color.White;
            UserName.Font = new Font("Segoe UI", 10F);
            UserName.Location = new Point(304, 122);
            UserName.Name = "UserName";
            UserName.Size = new Size(95, 23);
            UserName.TabIndex = 1;
            UserName.Text = "User Name";
            // 
            // textUserName
            // 
            textUserName.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            textUserName.Location = new Point(405, 118);
            textUserName.Name = "textUserName";
            textUserName.Size = new Size(125, 27);
            textUserName.TabIndex = 0;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(panel1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Form1";
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Panel panel1;
        private Button button1;
        private Label password;
        private TextBox textpassword;
        private Label UserName;
        private TextBox textUserName;
        private Label label1;
        private Label label2;
    }
}
