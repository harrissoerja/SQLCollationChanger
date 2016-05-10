using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Reflection;
using System.IO;
using System.Globalization;
using System.Threading;

namespace AlterCollation
{
    public class MainForm : Form, IScriptExecuteCallback
    {


		#region Windows Form Designer generated code


		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
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
		private System.Windows.Forms.Button CancelButton;
		private System.Windows.Forms.Label labelNote2;
		private System.Windows.Forms.Label labelNote1;
		private System.Windows.Forms.ProgressBar ProgressBar;
		private System.Windows.Forms.RichTextBox ReportText;
		private System.Windows.Forms.Button RunButton;
		private System.Windows.Forms.CheckBox integratedE;
        private GroupBox groupBox1;
        private CheckBox SetSingleUserCheck;
        private CheckBox DropConstraintCheck;
        private Label label1;
        private ComboBox AuthComboBox;
        private ComboBox LanguageCombobox;
        private Label languageL;
        private ComboBox CollationComboBox;
        private Label collationL;
        private Label databaseL;
        private Label passwordL;
        private Label userIdL;
        private Label serverL;
        private TextBox PasswordText;
        private TextBox UserNameText;
        private TextBox DatabaseNameText;
        private TextBox ServerInstanceText;
        private Button HelpButton;
        private Button TestConnButton;
        private CheckBox GenerateSQLOnlyText;
        private Button GenerateSQLbutton;
        private Button DirButton;
        private SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Timer cancelTimer;

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.labelNote2 = new System.Windows.Forms.Label();
            this.labelNote1 = new System.Windows.Forms.Label();
            this.ProgressBar = new System.Windows.Forms.ProgressBar();
            this.ReportText = new System.Windows.Forms.RichTextBox();
            this.integratedE = new System.Windows.Forms.CheckBox();
            this.cancelTimer = new System.Windows.Forms.Timer(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.GenerateSQLOnlyText = new System.Windows.Forms.CheckBox();
            this.SetSingleUserCheck = new System.Windows.Forms.CheckBox();
            this.DropConstraintCheck = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.DirButton = new System.Windows.Forms.Button();
            this.AuthComboBox = new System.Windows.Forms.ComboBox();
            this.LanguageCombobox = new System.Windows.Forms.ComboBox();
            this.languageL = new System.Windows.Forms.Label();
            this.CollationComboBox = new System.Windows.Forms.ComboBox();
            this.collationL = new System.Windows.Forms.Label();
            this.databaseL = new System.Windows.Forms.Label();
            this.passwordL = new System.Windows.Forms.Label();
            this.userIdL = new System.Windows.Forms.Label();
            this.serverL = new System.Windows.Forms.Label();
            this.PasswordText = new System.Windows.Forms.TextBox();
            this.UserNameText = new System.Windows.Forms.TextBox();
            this.DatabaseNameText = new System.Windows.Forms.TextBox();
            this.ServerInstanceText = new System.Windows.Forms.TextBox();
            this.GenerateSQLbutton = new System.Windows.Forms.Button();
            this.TestConnButton = new System.Windows.Forms.Button();
            this.HelpButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.RunButton = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelNote2
            // 
            this.labelNote2.Location = new System.Drawing.Point(170, 370);
            this.labelNote2.Name = "labelNote2";
            this.labelNote2.Size = new System.Drawing.Size(256, 56);
            this.labelNote2.TabIndex = 35;
            this.labelNote2.Text = "NOTE: Once the script has been executed you will see any error messages shown in " +
    "red in the window below directly after the SQL code that failed";
            this.labelNote2.Visible = false;
            // 
            // labelNote1
            // 
            this.labelNote1.Location = new System.Drawing.Point(170, 436);
            this.labelNote1.Name = "labelNote1";
            this.labelNote1.Size = new System.Drawing.Size(256, 80);
            this.labelNote1.TabIndex = 34;
            this.labelNote1.Text = resources.GetString("labelNote1.Text");
            this.labelNote1.Visible = false;
            // 
            // ProgressBar
            // 
            this.ProgressBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.ProgressBar.Location = new System.Drawing.Point(0, 559);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(615, 23);
            this.ProgressBar.TabIndex = 36;
            // 
            // ReportText
            // 
            this.ReportText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ReportText.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ReportText.DetectUrls = false;
            this.ReportText.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ReportText.ForeColor = System.Drawing.SystemColors.WindowText;
            this.ReportText.Location = new System.Drawing.Point(0, 304);
            this.ReportText.Name = "ReportText";
            this.ReportText.ReadOnly = true;
            this.ReportText.Size = new System.Drawing.Size(615, 249);
            this.ReportText.TabIndex = 37;
            this.ReportText.Text = "";
            // 
            // integratedE
            // 
            this.integratedE.Checked = true;
            this.integratedE.CheckState = System.Windows.Forms.CheckState.Checked;
            this.integratedE.Location = new System.Drawing.Point(173, 334);
            this.integratedE.Name = "integratedE";
            this.integratedE.Size = new System.Drawing.Size(184, 24);
            this.integratedE.TabIndex = 23;
            this.integratedE.Text = "Integrated Security";
            this.integratedE.Visible = false;
            this.integratedE.CheckedChanged += new System.EventHandler(this.integratedE_CheckedChanged);
            // 
            // cancelTimer
            // 
            this.cancelTimer.Interval = 5000;
            this.cancelTimer.Tick += new System.EventHandler(this.cancelTimer_Tick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.GenerateSQLOnlyText);
            this.groupBox1.Controls.Add(this.SetSingleUserCheck);
            this.groupBox1.Controls.Add(this.DropConstraintCheck);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.DirButton);
            this.groupBox1.Controls.Add(this.AuthComboBox);
            this.groupBox1.Controls.Add(this.LanguageCombobox);
            this.groupBox1.Controls.Add(this.languageL);
            this.groupBox1.Controls.Add(this.CollationComboBox);
            this.groupBox1.Controls.Add(this.collationL);
            this.groupBox1.Controls.Add(this.databaseL);
            this.groupBox1.Controls.Add(this.passwordL);
            this.groupBox1.Controls.Add(this.userIdL);
            this.groupBox1.Controls.Add(this.serverL);
            this.groupBox1.Controls.Add(this.PasswordText);
            this.groupBox1.Controls.Add(this.UserNameText);
            this.groupBox1.Controls.Add(this.DatabaseNameText);
            this.groupBox1.Controls.Add(this.ServerInstanceText);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(440, 286);
            this.groupBox1.TabIndex = 46;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Database and Connection Configuration";
            // 
            // GenerateSQLOnlyText
            // 
            this.GenerateSQLOnlyText.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.GenerateSQLOnlyText.Location = new System.Drawing.Point(122, 260);
            this.GenerateSQLOnlyText.Name = "GenerateSQLOnlyText";
            this.GenerateSQLOnlyText.Size = new System.Drawing.Size(256, 19);
            this.GenerateSQLOnlyText.TabIndex = 62;
            this.GenerateSQLOnlyText.Text = "Generate SQL File only";
            this.GenerateSQLOnlyText.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            // 
            // SetSingleUserCheck
            // 
            this.SetSingleUserCheck.Checked = true;
            this.SetSingleUserCheck.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SetSingleUserCheck.Location = new System.Drawing.Point(122, 213);
            this.SetSingleUserCheck.Name = "SetSingleUserCheck";
            this.SetSingleUserCheck.Size = new System.Drawing.Size(266, 17);
            this.SetSingleUserCheck.TabIndex = 61;
            this.SetSingleUserCheck.Text = "Set Database to single user mode (Recomended)";
            // 
            // DropConstraintCheck
            // 
            this.DropConstraintCheck.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.DropConstraintCheck.Location = new System.Drawing.Point(122, 235);
            this.DropConstraintCheck.Name = "DropConstraintCheck";
            this.DropConstraintCheck.Size = new System.Drawing.Size(256, 19);
            this.DropConstraintCheck.TabIndex = 60;
            this.DropConstraintCheck.Text = "Drop all constraints and keys first (optional)";
            this.DropConstraintCheck.TextAlign = System.Drawing.ContentAlignment.TopLeft;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 59;
            this.label1.Text = "Authentication";
            // 
            // DirButton
            // 
            this.DirButton.Location = new System.Drawing.Point(404, 28);
            this.DirButton.Name = "DirButton";
            this.DirButton.Size = new System.Drawing.Size(26, 20);
            this.DirButton.TabIndex = 58;
            this.DirButton.Text = "...";
            // 
            // AuthComboBox
            // 
            this.AuthComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.AuthComboBox.Items.AddRange(new object[] {
            "SQL Server Authentication",
            "Windows Authentication"});
            this.AuthComboBox.Location = new System.Drawing.Point(122, 54);
            this.AuthComboBox.Name = "AuthComboBox";
            this.AuthComboBox.Size = new System.Drawing.Size(279, 21);
            this.AuthComboBox.TabIndex = 57;
            // 
            // LanguageCombobox
            // 
            this.LanguageCombobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.LanguageCombobox.Location = new System.Drawing.Point(122, 186);
            this.LanguageCombobox.Name = "LanguageCombobox";
            this.LanguageCombobox.Size = new System.Drawing.Size(279, 21);
            this.LanguageCombobox.TabIndex = 56;
            // 
            // languageL
            // 
            this.languageL.AutoSize = true;
            this.languageL.Location = new System.Drawing.Point(16, 186);
            this.languageL.Name = "languageL";
            this.languageL.Size = new System.Drawing.Size(98, 13);
            this.languageL.TabIndex = 55;
            this.languageL.Text = "Full Text Language";
            // 
            // CollationComboBox
            // 
            this.CollationComboBox.Location = new System.Drawing.Point(122, 159);
            this.CollationComboBox.Name = "CollationComboBox";
            this.CollationComboBox.Size = new System.Drawing.Size(279, 21);
            this.CollationComboBox.TabIndex = 54;
            // 
            // collationL
            // 
            this.collationL.AutoSize = true;
            this.collationL.Location = new System.Drawing.Point(16, 159);
            this.collationL.Name = "collationL";
            this.collationL.Size = new System.Drawing.Size(81, 13);
            this.collationL.TabIndex = 53;
            this.collationL.Text = "Target Collation";
            // 
            // databaseL
            // 
            this.databaseL.AutoSize = true;
            this.databaseL.Location = new System.Drawing.Point(16, 133);
            this.databaseL.Name = "databaseL";
            this.databaseL.Size = new System.Drawing.Size(87, 13);
            this.databaseL.TabIndex = 51;
            this.databaseL.Text = "Database Name ";
            // 
            // passwordL
            // 
            this.passwordL.AutoSize = true;
            this.passwordL.Enabled = false;
            this.passwordL.Location = new System.Drawing.Point(16, 107);
            this.passwordL.Name = "passwordL";
            this.passwordL.Size = new System.Drawing.Size(53, 13);
            this.passwordL.TabIndex = 49;
            this.passwordL.Text = "Password";
            // 
            // userIdL
            // 
            this.userIdL.AutoSize = true;
            this.userIdL.Enabled = false;
            this.userIdL.Location = new System.Drawing.Point(16, 81);
            this.userIdL.Name = "userIdL";
            this.userIdL.Size = new System.Drawing.Size(60, 13);
            this.userIdL.TabIndex = 47;
            this.userIdL.Text = "User Name";
            // 
            // serverL
            // 
            this.serverL.AutoSize = true;
            this.serverL.Location = new System.Drawing.Point(16, 28);
            this.serverL.Name = "serverL";
            this.serverL.Size = new System.Drawing.Size(103, 13);
            this.serverL.TabIndex = 45;
            this.serverL.Text = "Server and Instance";
            // 
            // PasswordText
            // 
            this.PasswordText.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.PasswordText.Enabled = false;
            this.PasswordText.Location = new System.Drawing.Point(122, 107);
            this.PasswordText.Name = "PasswordText";
            this.PasswordText.Size = new System.Drawing.Size(279, 20);
            this.PasswordText.TabIndex = 50;
            // 
            // UserNameText
            // 
            this.UserNameText.Enabled = false;
            this.UserNameText.Location = new System.Drawing.Point(122, 81);
            this.UserNameText.Name = "UserNameText";
            this.UserNameText.Size = new System.Drawing.Size(279, 20);
            this.UserNameText.TabIndex = 48;
            // 
            // DatabaseNameText
            // 
            this.DatabaseNameText.Location = new System.Drawing.Point(122, 133);
            this.DatabaseNameText.Name = "DatabaseNameText";
            this.DatabaseNameText.Size = new System.Drawing.Size(279, 20);
            this.DatabaseNameText.TabIndex = 52;
            // 
            // ServerInstanceText
            // 
            this.ServerInstanceText.Location = new System.Drawing.Point(122, 28);
            this.ServerInstanceText.Name = "ServerInstanceText";
            this.ServerInstanceText.Size = new System.Drawing.Size(279, 20);
            this.ServerInstanceText.TabIndex = 46;
            // 
            // GenerateSQLbutton
            // 
            this.GenerateSQLbutton.Location = new System.Drawing.Point(458, 267);
            this.GenerateSQLbutton.Name = "GenerateSQLbutton";
            this.GenerateSQLbutton.Size = new System.Drawing.Size(112, 24);
            this.GenerateSQLbutton.TabIndex = 33;
            this.GenerateSQLbutton.Text = "Generate SQL File";
            this.GenerateSQLbutton.Click += new System.EventHandler(this.scriptB_Click);
            // 
            // TestConnButton
            // 
            this.TestConnButton.Image = global::CollationChanger.Resources._001_09;
            this.TestConnButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.TestConnButton.Location = new System.Drawing.Point(458, 29);
            this.TestConnButton.Name = "TestConnButton";
            this.TestConnButton.Size = new System.Drawing.Size(145, 31);
            this.TestConnButton.TabIndex = 62;
            this.TestConnButton.Text = "Test Connection";
            // 
            // HelpButton
            // 
            this.HelpButton.Image = global::CollationChanger.Resources._001_50;
            this.HelpButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.HelpButton.Location = new System.Drawing.Point(458, 139);
            this.HelpButton.Name = "HelpButton";
            this.HelpButton.Size = new System.Drawing.Size(145, 35);
            this.HelpButton.TabIndex = 47;
            this.HelpButton.Text = "Help";
            // 
            // CancelButton
            // 
            this.CancelButton.Image = global::CollationChanger.Resources._001_02;
            this.CancelButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.CancelButton.Location = new System.Drawing.Point(458, 100);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(145, 32);
            this.CancelButton.TabIndex = 41;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.Click += new System.EventHandler(this.cancelB_Click);
            // 
            // RunButton
            // 
            this.RunButton.Image = global::CollationChanger.Resources._001_25;
            this.RunButton.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.RunButton.Location = new System.Drawing.Point(458, 66);
            this.RunButton.Name = "RunButton";
            this.RunButton.Size = new System.Drawing.Size(145, 28);
            this.RunButton.TabIndex = 32;
            this.RunButton.Text = "&Run";
            this.RunButton.Click += new System.EventHandler(this.executeB_Click);
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(615, 582);
            this.Controls.Add(this.TestConnButton);
            this.Controls.Add(this.HelpButton);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.integratedE);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.labelNote2);
            this.Controls.Add(this.labelNote1);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.GenerateSQLbutton);
            this.Controls.Add(this.ReportText);
            this.Controls.Add(this.RunButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Atmosfer DB Collation Changer";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}

		#endregion

		


        private bool textLanguagePopulated;

        
        private bool collationPopulated;
        private ViewMode viewMode;
        private bool executeScriptOnly;
        private bool canceled;

        //private delegate ScriptStepCollection GenerateScript(IScriptExecuteCallback callback, string server, string userId, string password, string database, bool dropAllConstraints, string collation, FullTextLanguage language);
        

        private delegate void UpdateUICallback(int step);
        private delegate bool ScriptErrorCallback(ScriptStep step, Exception ex);
        private delegate void ScriptCompleteCallback(ScriptStepCollection script);
        private delegate void ScriptCompleteErrorCallback(Exception ex);
        private delegate void ExecuteCompleteCallback();
        
        private ScriptStepCollection executingScript;
        private int lastReportedStep;

        private Thread workerThread;
		private WorkerThreadArguments workerThreadArguments;

        private class WorkerThreadArguments
        {
            private bool scriptOnly;
            private IScriptExecuteCallback callback;
            private string server;
            private string userId;
            private string password;
            private string database;
            private bool dropAllConstraints;
            private string collation;
            private FullTextLanguage language;
			private bool setSingleUser;

            public WorkerThreadArguments(bool scriptOnly, IScriptExecuteCallback callback, string server, string userId, string password, string database, bool dropAllConstraints, string collation, FullTextLanguage language, bool setSingleUser)
            {
                this.scriptOnly = scriptOnly;
                this.callback = callback;
                this.server = server;
                this.userId = userId;
                this.password = password;
                this.database = database;
                this.dropAllConstraints = dropAllConstraints;
                this.collation = collation;
                this.language = language;
				this.setSingleUser =setSingleUser;
            }
            public bool ScriptOnly
            {
                get { return scriptOnly; }
            }
            public IScriptExecuteCallback Callback
            {
                get { return callback; }
            }
            public string Server
            {
                get { return server; }
            }
            public string UserId
            {
                get { return userId; }
            }
            public string Password
            {
                get { return password; }
            }
            public string Database
            {
                get { return database; }
            }
            public bool DropAllConstraints
            {
                get { return dropAllConstraints; }
            }
            public string Collation
            {
                get { return collation; }
            }
            public FullTextLanguage Language
            {
                get { return language; }
            }
			public bool SetSingleUser
			{
				get {return setSingleUser;}
			}

        }

        private enum ViewMode
        {
            Normal,
            Running,
            Aborting
        }

        public MainForm()
        {
            InitializeComponent();
        }


        private void Script()
        {
            if ((viewMode == ViewMode.Running))
            {
                MessageBox.Show("Already Running Something");
            }
            else
            {
                ViewModeSet(ViewMode.Running);
                ReportText.Clear();

                if (workerThread!=null)
                    throw new InvalidOperationException("Oops worker thread still running");

				//workerThread = new Thread(
				workerThreadArguments  =new WorkerThreadArguments(true,this, ServerInstanceText.Text, UserNameText.Text, PasswordText.Text, DatabaseNameText.Text, DropConstraintCheck.Checked, CollationComboBox.Text, LanguageCombobox.SelectedItem as FullTextLanguage, SetSingleUserCheck.Checked);
                workerThread = new Thread(new ThreadStart(ScriptThreadProc));
                workerThread.Start();


            }
        }

        private void ScriptThreadProc()
        {
            //WorkerThreadArguments arguments = (WorkerThreadArguments)threadArguments;
            CollationChanger collationChanger = new CollationChanger();
            ScriptStepCollection script = null;
            SqlConnection connection = null;
            try
            {
                script = collationChanger.GenerateScript(workerThreadArguments.Callback, workerThreadArguments.Server, workerThreadArguments.UserId, workerThreadArguments.Password, workerThreadArguments.Database, workerThreadArguments.DropAllConstraints, workerThreadArguments.Collation, workerThreadArguments.Language, workerThreadArguments.SetSingleUser);
				if (script!=null)
				{
					if (workerThreadArguments.ScriptOnly)
					{
						BeginInvoke(new ScriptCompleteCallback(ScriptComplete), new object[] {script});
					}
					else
					{
						connection = new SqlConnection(Utils.ConnectionString(workerThreadArguments.Server, workerThreadArguments.UserId, workerThreadArguments.Password));
						connection.Open();
						script.Execute(connection, workerThreadArguments.Callback);
						BeginInvoke(new ExecuteCompleteCallback(ExecuteComplete));
					}
				}
				else
				{
					BeginInvoke(new ScriptCompleteCallback(ScriptComplete), new object[] { null});
				}


            }
            catch (ThreadAbortException) { throw; }
            catch (Exception ex)
            {
				BeginInvoke(new ScriptCompleteErrorCallback(ScriptComplete), new object[] { ex});
            }
            finally
            {
                if (connection != null)
                    connection.Dispose();
            }

            lock (this)
            {
                workerThread = null;
            }
            
        }

        private void ExecuteComplete()
        {
            if (!canceled)
                MessageBox.Show("Execution Complete", this.Text);
            ProgressBar.Value = ProgressBar.Maximum;
            ViewModeSet(ViewMode.Normal);
        }
        
        /// <summary>
        /// callback from worker thread when something goes wrong
        /// </summary>
        /// <param name="ex"></param>
        private void ScriptComplete(Exception ex)
        {
            MessageBox.Show(ex.Message, this.Text);
            ProgressBar.Value = ProgressBar.Maximum;
            ViewModeSet(ViewMode.Normal);
        }
        
        /// <summary>
        /// callback from worker thread when it went OK
        /// </summary>
        /// <param name="script"></param>
        private void ScriptComplete(ScriptStepCollection script)
        {
            if (script!=null)
                WriteScriptToWindow(script);
            ProgressBar.Value = ProgressBar.Maximum;
            ViewModeSet(ViewMode.Normal);
        }


        private void WriteScriptToWindow(ScriptStepCollection script)
        {
            ReportText.Clear();
            foreach (ScriptStep step in script)
            {
                ReportText.AppendText(step.CommandText);
                ReportText.AppendText("\n\nGO\n\n");

            }
        }

        private void ViewModeSet(ViewMode viewMode)
        {
            this.viewMode = viewMode;
            ServerInstanceText.Enabled = viewMode == ViewMode.Normal;
            integratedE.Enabled = viewMode == ViewMode.Normal;
            UserNameText.Enabled = viewMode == ViewMode.Normal && !integratedE.Checked;
            PasswordText.Enabled = viewMode == ViewMode.Normal && !integratedE.Checked;
            DatabaseNameText.Enabled = viewMode == ViewMode.Normal;
            CollationComboBox.Enabled = viewMode == ViewMode.Normal;
            LanguageCombobox.Enabled = viewMode == ViewMode.Normal;
            GenerateSQLbutton.Enabled = viewMode == ViewMode.Normal;
            RunButton.Enabled = viewMode == ViewMode.Normal;
            DropConstraintCheck.Enabled = viewMode == ViewMode.Normal;
            CancelButton.Enabled = viewMode == ViewMode.Running;
			SetSingleUserCheck.Enabled = viewMode == ViewMode.Normal;
            cancelTimer.Enabled = viewMode == ViewMode.Aborting;

            switch (viewMode)
            {
                case ViewMode.Running:
                    canceled = false;
                    break;
                case ViewMode.Aborting:
                    canceled = true;
                    break;
                case ViewMode.Normal:
                    break;
            }
        }



        

        //private void ExecuteDo()
        //{
        //    canceled = false;
        //    ScriptRunner scriptRunner = new ScriptRunner();

        //    scriptRunner.Execute(serverE.Text, userIdE.Text, passwordE.Text,databaseE.Text, dropAllE.Checked, collationE.Text, (FullTextLanguage)languageE.SelectedItem);
        //    ExecuteComplete();

        //}


        protected override void OnLoad(System.EventArgs e)
        {
            base.OnLoad(e);
            ViewModeSet(ViewMode.Normal);

        }

        private void executeB_Click(object sender, EventArgs e)
        {

            if ((viewMode == ViewMode.Running))
            {
                MessageBox.Show("Already Running Something");
            }
            else
            {
                if (MessageBox.Show("This program will now execute a script to alter the collation of your database.  This may take a long time and may result in data loss.  Please ensure that all your data is backed up.\n\nExclusive database access is required to complete the process so before running use the sp_who command to verify that there are no users connected to your database.", "Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2) == DialogResult.OK)
                {
                    ViewModeSet(ViewMode.Running);
                    ReportText.Clear();

                    if (workerThread != null)
                        throw new InvalidOperationException("Oops worker thread still running");

                    workerThreadArguments = new WorkerThreadArguments(false, this, ServerInstanceText.Text, UserNameText.Text, PasswordText.Text, DatabaseNameText.Text, DropConstraintCheck.Checked, CollationComboBox.Text, LanguageCombobox.SelectedItem as FullTextLanguage, SetSingleUserCheck.Checked);
                    workerThread = new Thread(new ThreadStart(ScriptThreadProc));
                    workerThread.Start();
                }
                else {

                }
                
            }

        }

        private void integratedE_CheckedChanged(object sender, EventArgs e)
        {
            if (integratedE.Checked)
            {
                UserNameText.Text = string.Empty;
                PasswordText.Text = string.Empty;
            }

            UserNameText.Enabled = !integratedE.Checked;
            userIdL.Enabled = !integratedE.Checked;
            PasswordText.Enabled = !integratedE.Checked;
            passwordL.Enabled = !integratedE.Checked;

        }

        private void scriptB_Click(object sender, EventArgs e)
        {
            Script();
        }

        private void languageE_DropDown(object sender, EventArgs e)
        {
            if (!textLanguagePopulated)
            {
                textLanguagePopulated = true;
                SqlConnection con = new SqlConnection();

                try
                {

                    SqlCommand cmd;
                    int ixName;
                    int ixLcid;
                    SqlDataReader dr;
                    Version serverVersion;
                    con.ConnectionString = Utils.ConnectionString(ServerInstanceText.Text, UserNameText.Text, PasswordText.Text);
                    con.Open();
                    serverVersion = new Version(con.ServerVersion);

                    cmd = con.CreateCommand();


                    if (serverVersion.Major>=9)
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "select * from sys.fulltext_languages";
                    }
                    else
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "master..xp_MSFullText";
                    }

                    dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    
                    
                    if (serverVersion.Major>=9)
                    {
                        ixName = dr.GetOrdinal("name");
                    }
                    else
                    {
                        ixName = dr.GetOrdinal("Language");
                    }

                    ixLcid = dr.GetOrdinal("LCID");

                    LanguageCombobox.Items.Clear();
                    LanguageCombobox.Items.Add(new FullTextLanguage("<Unchanged>", int.MinValue));
                    while (dr.Read())
                    {
                        LanguageCombobox.Items.Add(new FullTextLanguage(dr.GetString(ixName), dr.GetInt32(ixLcid)));
                    }
                }

                catch (SqlException)
                {
                    textLanguagePopulated = false;
                }
                catch (Exception)
                {
                    textLanguagePopulated = false;
                    throw;
                }
                finally
                {
                    con.Dispose();
                }
            }
        }

        private void cancelB_Click(object sender, EventArgs e)
        {
            ViewModeSet(ViewMode.Aborting);
        }

        private void serverE_TextChanged(object sender, EventArgs e)
        {
            collationPopulated = false;
            textLanguagePopulated = false;
            CollationComboBox.Items.Clear();

            LanguageCombobox.Items.Clear();
            LanguageCombobox.Items.Add(new FullTextLanguage("<Unchanged>", int.MinValue));
        }

        private void collationE_DropDown(object sender, EventArgs e)
        {
            //populate the list

            if (!collationPopulated)
            {
                collationPopulated = true;
                SqlConnection con = new SqlConnection();

                try
                {

                    SqlCommand cmd;
                    int ixName;
                    SqlDataReader dr;

                    con.ConnectionString = Utils.ConnectionString(ServerInstanceText.Text, UserNameText.Text, PasswordText.Text);
                    con.Open();
                    cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select name from ::fn_helpCollations()";
                    dr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    ixName = dr.GetOrdinal("name");
                    CollationComboBox.Items.Clear();
                    while (dr.Read())
                    {
                        CollationComboBox.Items.Add(dr.GetString(ixName));
                    }
                }
                catch (SqlException)
                {
                    collationPopulated = false;
                }
                catch (Exception)
                {
                    collationPopulated = false;
                    throw;
                }
                finally
                {
                    con.Dispose();
                }
            }
        }


        private void UpdateUI(int step)
        {
			lock(this)
			{
				if (lastReportedStep == -1)
				{
					ReportText.Clear();
					lastReportedStep =0;
				}
				ScriptStep scriptItem;
				for (int index = lastReportedStep; index < step; index++)
				{
					scriptItem = executingScript[index];
					ReportText.AppendText(scriptItem.CommandText);
					ReportText.AppendText("\n\nGO\n\n");

				}
				ProgressBar.Maximum = executingScript.Count;
				ProgressBar.Value = step;
				lastReportedStep = step;
			}
        }


        private bool ScriptError(ScriptStep step, Exception exception)
        {
            SqlException ex = exception as SqlException;
            if (ex == null)
            {
                MessageBox.Show("Unexpeced error");
                return false;
            }

            for (int stepIndex = 0; stepIndex < executingScript.Count; stepIndex++)
            {
                if (object.ReferenceEquals(executingScript[stepIndex], step))
                {
                    UpdateUI(stepIndex+1);
                    break;
                }
            }


            foreach (SqlError e in ex.Errors)
            {
                if (e.Class >= 11)
                {
                    ReportText.SelectionColor = Color.Red;
                }
                else
                {
                    ReportText.SelectionColor = Color.Black;
                }
                ReportText.AppendText(string.Format("{0} - {1}", e.Number, e.Message));
                ReportText.AppendText("\n");
            }

            if (ex.Class >= 11)
            {
                string commandText = step.CommandText;
                if (commandText.Length > 1000)
                    commandText = commandText.Substring(0, 1000) + "......";

                string message = string.Format("An error occured while executing the SQL Statement\n\n '{0}'\n\n{1}\n\nDo you which to continue running the script anyway?", commandText, ex.Message);
                if (!(MessageBox.Show(this,message, "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes))
                {
                    canceled = true;
                    ReportText.SelectionColor = Color.Red;
                    ReportText.AppendText("Canceled");
                    ReportText.AppendText("\n");
                }
            }


            return !canceled;
        }

        #region IScriptExecuteCallback Members

        bool IScriptExecuteCallback.Error(ScriptStep step, Exception exception)
        {
			return (bool) this.Invoke(new ScriptErrorCallback(ScriptError), new object [] {step, exception});
        }

        bool IScriptExecuteCallback.Progress(ScriptStep script, int step, int stepMax)
        {
            if (!canceled)
				this.BeginInvoke(new UpdateUICallback(UpdateUI), new object[] {step});
            return !canceled;
        }

        void IScriptExecuteCallback.ExecutionStarting(ScriptStepCollection script)
        {
			lock(this)
			{
				executingScript = script;
				lastReportedStep = -1;
			}
        }
        #endregion

        private void cancelTimer_Tick(object sender, EventArgs e)
        {
            lock(this)
            {
                if (workerThread != null)
                    workerThread.Abort();
                workerThread = null;
            }
            ProgressBar.Value = ProgressBar.Maximum;
            ViewModeSet(ViewMode.Normal);
            
        }

        private void collationE_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            AuthComboBox.SelectedIndex = 1;
        }
    }
}



