namespace Auto_Email
{
    partial class MainPage
    {
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.gvEmailList = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comboxCaseStatus = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnEmailLogin = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageEmailList = new System.Windows.Forms.TabPage();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.btnEmailSearch = new System.Windows.Forms.Button();
            this.dtPickerDateTo = new System.Windows.Forms.DateTimePicker();
            this.dtPickerDateFrom = new System.Windows.Forms.DateTimePicker();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tBLastTimeRun = new System.Windows.Forms.TextBox();
            this.tabPageRejectPattern = new System.Windows.Forms.TabPage();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.dgvActionRequired = new System.Windows.Forms.DataGridView();
            this.dgvRejectReasonPattern = new System.Windows.Forms.DataGridView();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.txtActionRequired = new System.Windows.Forms.TextBox();
            this.txtActionCategory = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.txtReasonPattern = new System.Windows.Forms.TextBox();
            this.txtRejectCategory = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tabPageEmailLogin = new System.Windows.Forms.TabPage();
            this.btnDelete = new System.Windows.Forms.Button();
            this.dgvEmailLogin = new System.Windows.Forms.DataGridView();
            this.btnClearEmailLogin = new System.Windows.Forms.Button();
            this.btnSaveEmailLogin = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.lbl_LoginID = new System.Windows.Forms.Label();
            this.tbLoginPortNo = new System.Windows.Forms.TextBox();
            this.tbLoginEmailAddress = new System.Windows.Forms.TextBox();
            this.tbLoginPassword = new System.Windows.Forms.TextBox();
            this.tbLoginUsername = new System.Windows.Forms.TextBox();
            this.tbLoginHostname = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.gvEmailList)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPageEmailList.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.tabPageRejectPattern.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvActionRequired)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRejectReasonPattern)).BeginInit();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.tabPageEmailLogin.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmailLogin)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // gvEmailList
            // 
            this.gvEmailList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvEmailList.Location = new System.Drawing.Point(6, 75);
            this.gvEmailList.Name = "gvEmailList";
            this.gvEmailList.Size = new System.Drawing.Size(1095, 430);
            this.gvEmailList.TabIndex = 0;
            this.gvEmailList.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gvEmailList_CellClick);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.groupBox1.Controls.Add(this.comboxCaseStatus);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(608, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(265, 53);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filter";
            this.groupBox1.Visible = false;
            // 
            // comboxCaseStatus
            // 
            this.comboxCaseStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboxCaseStatus.Items.AddRange(new object[] {
            "All Case Status",
            "Open Case Status",
            "Close Case Status"});
            this.comboxCaseStatus.Location = new System.Drawing.Point(101, 19);
            this.comboxCaseStatus.Name = "comboxCaseStatus";
            this.comboxCaseStatus.Size = new System.Drawing.Size(121, 21);
            this.comboxCaseStatus.TabIndex = 2;
            this.comboxCaseStatus.SelectedIndexChanged += new System.EventHandler(this.comboxCaseStatus_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(22, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Case Status : ";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnEmailLogin);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Location = new System.Drawing.Point(12, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(89, 262);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Menu";
            // 
            // btnEmailLogin
            // 
            this.btnEmailLogin.Location = new System.Drawing.Point(6, 151);
            this.btnEmailLogin.Name = "btnEmailLogin";
            this.btnEmailLogin.Size = new System.Drawing.Size(75, 52);
            this.btnEmailLogin.TabIndex = 2;
            this.btnEmailLogin.Text = "Email Address Login";
            this.btnEmailLogin.UseVisualStyleBackColor = true;
            this.btnEmailLogin.Click += new System.EventHandler(this.btnEmailLogin_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(6, 93);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 52);
            this.button2.TabIndex = 1;
            this.button2.Text = "Reject Email Pattern";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(6, 35);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 52);
            this.button1.TabIndex = 0;
            this.button1.Text = "Rejected Email List";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPageEmailList);
            this.tabControl1.Controls.Add(this.tabPageRejectPattern);
            this.tabControl1.Controls.Add(this.tabPageEmailLogin);
            this.tabControl1.Location = new System.Drawing.Point(107, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1115, 537);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPageEmailList
            // 
            this.tabPageEmailList.Controls.Add(this.groupBox6);
            this.tabPageEmailList.Controls.Add(this.label6);
            this.tabPageEmailList.Controls.Add(this.tBLastTimeRun);
            this.tabPageEmailList.Controls.Add(this.groupBox1);
            this.tabPageEmailList.Controls.Add(this.gvEmailList);
            this.tabPageEmailList.Location = new System.Drawing.Point(4, 22);
            this.tabPageEmailList.Name = "tabPageEmailList";
            this.tabPageEmailList.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageEmailList.Size = new System.Drawing.Size(1107, 511);
            this.tabPageEmailList.TabIndex = 0;
            this.tabPageEmailList.Text = "Rejected Email List";
            this.tabPageEmailList.UseVisualStyleBackColor = true;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.btnEmailSearch);
            this.groupBox6.Controls.Add(this.dtPickerDateTo);
            this.groupBox6.Controls.Add(this.dtPickerDateFrom);
            this.groupBox6.Controls.Add(this.label15);
            this.groupBox6.Controls.Add(this.label14);
            this.groupBox6.Location = new System.Drawing.Point(6, 12);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(596, 57);
            this.groupBox6.TabIndex = 4;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Filter By:";
            this.groupBox6.Visible = false;
            // 
            // btnEmailSearch
            // 
            this.btnEmailSearch.Location = new System.Drawing.Point(431, 19);
            this.btnEmailSearch.Name = "btnEmailSearch";
            this.btnEmailSearch.Size = new System.Drawing.Size(75, 23);
            this.btnEmailSearch.TabIndex = 6;
            this.btnEmailSearch.Text = "Search";
            this.btnEmailSearch.UseVisualStyleBackColor = true;
            this.btnEmailSearch.Click += new System.EventHandler(this.btnEmailSearch_Click);
            // 
            // dtPickerDateTo
            // 
            this.dtPickerDateTo.AllowDrop = true;
            this.dtPickerDateTo.CustomFormat = "dd/MMM/yyyy";
            this.dtPickerDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtPickerDateTo.Location = new System.Drawing.Point(278, 22);
            this.dtPickerDateTo.Name = "dtPickerDateTo";
            this.dtPickerDateTo.Size = new System.Drawing.Size(107, 20);
            this.dtPickerDateTo.TabIndex = 5;
            // 
            // dtPickerDateFrom
            // 
            this.dtPickerDateFrom.AllowDrop = true;
            this.dtPickerDateFrom.CustomFormat = "dd/MMM/yyyy";
            this.dtPickerDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtPickerDateFrom.Location = new System.Drawing.Point(74, 22);
            this.dtPickerDateFrom.Name = "dtPickerDateFrom";
            this.dtPickerDateFrom.Size = new System.Drawing.Size(107, 20);
            this.dtPickerDateFrom.TabIndex = 4;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(220, 26);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(52, 13);
            this.label15.TabIndex = 1;
            this.label15.Text = "Date To: ";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(6, 26);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(62, 13);
            this.label14.TabIndex = 0;
            this.label14.Text = "Date From: ";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(879, 34);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(82, 13);
            this.label6.TabIndex = 3;
            this.label6.Text = "Last Time Run :";
            // 
            // tBLastTimeRun
            // 
            this.tBLastTimeRun.Location = new System.Drawing.Point(967, 31);
            this.tBLastTimeRun.Name = "tBLastTimeRun";
            this.tBLastTimeRun.Size = new System.Drawing.Size(134, 20);
            this.tBLastTimeRun.TabIndex = 2;
            // 
            // tabPageRejectPattern
            // 
            this.tabPageRejectPattern.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.tabPageRejectPattern.Controls.Add(this.label13);
            this.tabPageRejectPattern.Controls.Add(this.label12);
            this.tabPageRejectPattern.Controls.Add(this.dgvActionRequired);
            this.tabPageRejectPattern.Controls.Add(this.dgvRejectReasonPattern);
            this.tabPageRejectPattern.Controls.Add(this.btnClear);
            this.tabPageRejectPattern.Controls.Add(this.btnSave);
            this.tabPageRejectPattern.Controls.Add(this.groupBox4);
            this.tabPageRejectPattern.Controls.Add(this.groupBox3);
            this.tabPageRejectPattern.Location = new System.Drawing.Point(4, 22);
            this.tabPageRejectPattern.Name = "tabPageRejectPattern";
            this.tabPageRejectPattern.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageRejectPattern.Size = new System.Drawing.Size(1107, 511);
            this.tabPageRejectPattern.TabIndex = 1;
            this.tabPageRejectPattern.Text = "Email Reject Pattern";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(560, 154);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(113, 13);
            this.label13.TabIndex = 10;
            this.label13.Text = "Action Required Table";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(22, 154);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(157, 13);
            this.label12.TabIndex = 9;
            this.label12.Text = "Rejected Reason Pattern Table";
            // 
            // dgvActionRequired
            // 
            this.dgvActionRequired.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvActionRequired.Location = new System.Drawing.Point(554, 170);
            this.dgvActionRequired.Name = "dgvActionRequired";
            this.dgvActionRequired.Size = new System.Drawing.Size(500, 335);
            this.dgvActionRequired.TabIndex = 8;
            // 
            // dgvRejectReasonPattern
            // 
            this.dgvRejectReasonPattern.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvRejectReasonPattern.Location = new System.Drawing.Point(16, 170);
            this.dgvRejectReasonPattern.Name = "dgvRejectReasonPattern";
            this.dgvRejectReasonPattern.Size = new System.Drawing.Size(500, 335);
            this.dgvRejectReasonPattern.TabIndex = 7;
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(1006, 62);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(75, 23);
            this.btnClear.TabIndex = 6;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(1006, 33);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.txtActionRequired);
            this.groupBox4.Controls.Add(this.txtActionCategory);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Location = new System.Drawing.Point(554, 12);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(431, 118);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Action Detail";
            // 
            // txtActionRequired
            // 
            this.txtActionRequired.Location = new System.Drawing.Point(141, 49);
            this.txtActionRequired.Multiline = true;
            this.txtActionRequired.Name = "txtActionRequired";
            this.txtActionRequired.Size = new System.Drawing.Size(263, 53);
            this.txtActionRequired.TabIndex = 2;
            // 
            // txtActionCategory
            // 
            this.txtActionCategory.Location = new System.Drawing.Point(141, 23);
            this.txtActionCategory.Name = "txtActionCategory";
            this.txtActionCategory.ReadOnly = true;
            this.txtActionCategory.Size = new System.Drawing.Size(263, 20);
            this.txtActionCategory.TabIndex = 1;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(46, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "Action Required :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(129, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Reject Reason Category :";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.txtReasonPattern);
            this.groupBox3.Controls.Add(this.txtRejectCategory);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(16, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(431, 118);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Reject Reason Detail";
            // 
            // txtReasonPattern
            // 
            this.txtReasonPattern.Location = new System.Drawing.Point(141, 49);
            this.txtReasonPattern.Multiline = true;
            this.txtReasonPattern.Name = "txtReasonPattern";
            this.txtReasonPattern.Size = new System.Drawing.Size(263, 53);
            this.txtReasonPattern.TabIndex = 2;
            // 
            // txtRejectCategory
            // 
            this.txtRejectCategory.Location = new System.Drawing.Point(141, 23);
            this.txtRejectCategory.Name = "txtRejectCategory";
            this.txtRejectCategory.Size = new System.Drawing.Size(263, 20);
            this.txtRejectCategory.TabIndex = 1;
            this.txtRejectCategory.Leave += new System.EventHandler(this.txtRejectCategory_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(121, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Reject Reason Pattern :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(129, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Reject Reason Category :";
            // 
            // tabPageEmailLogin
            // 
            this.tabPageEmailLogin.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.tabPageEmailLogin.Controls.Add(this.btnDelete);
            this.tabPageEmailLogin.Controls.Add(this.dgvEmailLogin);
            this.tabPageEmailLogin.Controls.Add(this.btnClearEmailLogin);
            this.tabPageEmailLogin.Controls.Add(this.btnSaveEmailLogin);
            this.tabPageEmailLogin.Controls.Add(this.groupBox5);
            this.tabPageEmailLogin.Location = new System.Drawing.Point(4, 22);
            this.tabPageEmailLogin.Name = "tabPageEmailLogin";
            this.tabPageEmailLogin.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageEmailLogin.Size = new System.Drawing.Size(1107, 511);
            this.tabPageEmailLogin.TabIndex = 2;
            this.tabPageEmailLogin.Text = "Email Address Login";
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(418, 397);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 6;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // dgvEmailLogin
            // 
            this.dgvEmailLogin.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvEmailLogin.Location = new System.Drawing.Point(19, 12);
            this.dgvEmailLogin.Name = "dgvEmailLogin";
            this.dgvEmailLogin.Size = new System.Drawing.Size(1068, 304);
            this.dgvEmailLogin.TabIndex = 5;
            this.dgvEmailLogin.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvEmailLogin_CellClick);
            this.dgvEmailLogin.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvEmailLogin_CellFormatting);
            // 
            // btnClearEmailLogin
            // 
            this.btnClearEmailLogin.Location = new System.Drawing.Point(418, 372);
            this.btnClearEmailLogin.Name = "btnClearEmailLogin";
            this.btnClearEmailLogin.Size = new System.Drawing.Size(75, 23);
            this.btnClearEmailLogin.TabIndex = 2;
            this.btnClearEmailLogin.Text = "Clear";
            this.btnClearEmailLogin.UseVisualStyleBackColor = true;
            this.btnClearEmailLogin.Click += new System.EventHandler(this.btnClearEmailLogin_Click_1);
            // 
            // btnSaveEmailLogin
            // 
            this.btnSaveEmailLogin.Location = new System.Drawing.Point(418, 346);
            this.btnSaveEmailLogin.Name = "btnSaveEmailLogin";
            this.btnSaveEmailLogin.Size = new System.Drawing.Size(75, 23);
            this.btnSaveEmailLogin.TabIndex = 1;
            this.btnSaveEmailLogin.Text = "Save";
            this.btnSaveEmailLogin.UseVisualStyleBackColor = true;
            this.btnSaveEmailLogin.Click += new System.EventHandler(this.btnSaveEmailLogin_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.lbl_LoginID);
            this.groupBox5.Controls.Add(this.tbLoginPortNo);
            this.groupBox5.Controls.Add(this.tbLoginEmailAddress);
            this.groupBox5.Controls.Add(this.tbLoginPassword);
            this.groupBox5.Controls.Add(this.tbLoginUsername);
            this.groupBox5.Controls.Add(this.tbLoginHostname);
            this.groupBox5.Controls.Add(this.label11);
            this.groupBox5.Controls.Add(this.label10);
            this.groupBox5.Controls.Add(this.label9);
            this.groupBox5.Controls.Add(this.label8);
            this.groupBox5.Controls.Add(this.label7);
            this.groupBox5.Location = new System.Drawing.Point(19, 322);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(386, 183);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Email Login Detail";
            // 
            // lbl_LoginID
            // 
            this.lbl_LoginID.AutoSize = true;
            this.lbl_LoginID.Location = new System.Drawing.Point(97, 157);
            this.lbl_LoginID.Name = "lbl_LoginID";
            this.lbl_LoginID.Size = new System.Drawing.Size(0, 13);
            this.lbl_LoginID.TabIndex = 10;
            this.lbl_LoginID.Visible = false;
            // 
            // tbLoginPortNo
            // 
            this.tbLoginPortNo.Location = new System.Drawing.Point(97, 130);
            this.tbLoginPortNo.Name = "tbLoginPortNo";
            this.tbLoginPortNo.Size = new System.Drawing.Size(91, 20);
            this.tbLoginPortNo.TabIndex = 9;
            // 
            // tbLoginEmailAddress
            // 
            this.tbLoginEmailAddress.Location = new System.Drawing.Point(97, 52);
            this.tbLoginEmailAddress.Name = "tbLoginEmailAddress";
            this.tbLoginEmailAddress.Size = new System.Drawing.Size(270, 20);
            this.tbLoginEmailAddress.TabIndex = 6;
            // 
            // tbLoginPassword
            // 
            this.tbLoginPassword.Location = new System.Drawing.Point(97, 104);
            this.tbLoginPassword.Name = "tbLoginPassword";
            this.tbLoginPassword.PasswordChar = '*';
            this.tbLoginPassword.Size = new System.Drawing.Size(150, 20);
            this.tbLoginPassword.TabIndex = 8;
            // 
            // tbLoginUsername
            // 
            this.tbLoginUsername.Location = new System.Drawing.Point(97, 78);
            this.tbLoginUsername.Name = "tbLoginUsername";
            this.tbLoginUsername.Size = new System.Drawing.Size(150, 20);
            this.tbLoginUsername.TabIndex = 7;
            // 
            // tbLoginHostname
            // 
            this.tbLoginHostname.Location = new System.Drawing.Point(97, 26);
            this.tbLoginHostname.Name = "tbLoginHostname";
            this.tbLoginHostname.Size = new System.Drawing.Size(270, 20);
            this.tbLoginHostname.TabIndex = 5;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(39, 133);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(52, 13);
            this.label11.TabIndex = 4;
            this.label11.Text = "Port No : ";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(9, 55);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(82, 13);
            this.label10.TabIndex = 3;
            this.label10.Text = "Email Address : ";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(27, 107);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(62, 13);
            this.label9.TabIndex = 2;
            this.label9.Text = "Password : ";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(27, 81);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(64, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "Username : ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(27, 29);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 13);
            this.label7.TabIndex = 0;
            this.label7.Text = "Hostname : ";
            // 
            // timer1
            // 
            this.timer1.Interval = 180000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // MainPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1234, 562);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.tabControl1);
            this.Name = "MainPage";
            this.ShowIcon = false;
            this.Text = "Main Page";
            this.Load += new System.EventHandler(this.MainPage_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gvEmailList)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPageEmailList.ResumeLayout(false);
            this.tabPageEmailList.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.tabPageRejectPattern.ResumeLayout(false);
            this.tabPageRejectPattern.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvActionRequired)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRejectReasonPattern)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.tabPageEmailLogin.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvEmailLogin)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView gvEmailList;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox comboxCaseStatus;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageEmailList;
        private System.Windows.Forms.TabPage tabPageRejectPattern;
        private System.Windows.Forms.TextBox txtRejectCategory;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtReasonPattern;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox txtActionRequired;
        private System.Windows.Forms.TextBox txtActionCategory;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tBLastTimeRun;
        private System.Windows.Forms.TabPage tabPageEmailLogin;
        private System.Windows.Forms.Button btnEmailLogin;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button btnClearEmailLogin;
        private System.Windows.Forms.Button btnSaveEmailLogin;
        private System.Windows.Forms.TextBox tbLoginPortNo;
        private System.Windows.Forms.TextBox tbLoginEmailAddress;
        private System.Windows.Forms.TextBox tbLoginPassword;
        private System.Windows.Forms.TextBox tbLoginUsername;
        private System.Windows.Forms.TextBox tbLoginHostname;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DataGridView dgvEmailLogin;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Label lbl_LoginID;
        private System.Windows.Forms.DataGridView dgvRejectReasonPattern;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DataGridView dgvActionRequired;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.DateTimePicker dtPickerDateFrom;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button btnEmailSearch;
        private System.Windows.Forms.DateTimePicker dtPickerDateTo;
    }
}

