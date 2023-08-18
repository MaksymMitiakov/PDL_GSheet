namespace GSheets
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
			System.Windows.Forms.Label label1;
			System.Windows.Forms.Label label2;
			System.Windows.Forms.Label label3;
			System.Windows.Forms.Label label4;
			System.Windows.Forms.Label label5;
			System.Windows.Forms.Label label6;
			this.dataGrid = new System.Windows.Forms.DataGridView();
			this.No = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.FirstName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.LastName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Title = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.CompanyName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.PersonLinkedinUrl = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Email = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Website = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.mobile_phone = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.btnLoad = new System.Windows.Forms.Button();
			this.txtSpreadSheetId = new System.Windows.Forms.TextBox();
			this.txtSheetName = new System.Windows.Forms.TextBox();
			this.txtColumnStart = new System.Windows.Forms.TextBox();
			this.txtColumnEnd = new System.Windows.Forms.TextBox();
			this.txtFirstRow = new System.Windows.Forms.TextBox();
			this.btnPopulate = new System.Windows.Forms.Button();
			this.txtPDLKey = new System.Windows.Forms.TextBox();
			this.chkValidMobilePhone = new System.Windows.Forms.CheckBox();
			this.chkValidEmail = new System.Windows.Forms.CheckBox();
			label1 = new System.Windows.Forms.Label();
			label2 = new System.Windows.Forms.Label();
			label3 = new System.Windows.Forms.Label();
			label4 = new System.Windows.Forms.Label();
			label5 = new System.Windows.Forms.Label();
			label6 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dataGrid)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			label1.AutoSize = true;
			label1.Location = new System.Drawing.Point(12, 41);
			label1.Name = "label1";
			label1.Size = new System.Drawing.Size(89, 15);
			label1.TabIndex = 2;
			label1.Text = "SpreadSheet ID:";
			// 
			// label2
			// 
			label2.AutoSize = true;
			label2.Location = new System.Drawing.Point(27, 70);
			label2.Name = "label2";
			label2.Size = new System.Drawing.Size(74, 15);
			label2.TabIndex = 4;
			label2.Text = "Sheet Name:";
			// 
			// label3
			// 
			label3.AutoSize = true;
			label3.Location = new System.Drawing.Point(266, 70);
			label3.Name = "label3";
			label3.Size = new System.Drawing.Size(89, 15);
			label3.TabIndex = 6;
			label3.Text = "Column Range:";
			// 
			// label4
			// 
			label4.AutoSize = true;
			label4.Location = new System.Drawing.Point(390, 70);
			label4.Name = "label4";
			label4.Size = new System.Drawing.Size(12, 15);
			label4.TabIndex = 8;
			label4.Text = "-";
			// 
			// label5
			// 
			label5.AutoSize = true;
			label5.Location = new System.Drawing.Point(454, 70);
			label5.Name = "label5";
			label5.Size = new System.Drawing.Size(58, 15);
			label5.TabIndex = 10;
			label5.Text = "First Row:";
			// 
			// label6
			// 
			label6.AutoSize = true;
			label6.Location = new System.Drawing.Point(27, 12);
			label6.Name = "label6";
			label6.Size = new System.Drawing.Size(74, 15);
			label6.TabIndex = 13;
			label6.Text = "PDL API Key:";
			// 
			// dataGrid
			// 
			this.dataGrid.AllowUserToAddRows = false;
			this.dataGrid.AllowUserToDeleteRows = false;
			this.dataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.dataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.No,
            this.FirstName,
            this.LastName,
            this.Title,
            this.CompanyName,
            this.PersonLinkedinUrl,
            this.Email,
            this.Website,
            this.mobile_phone});
			this.dataGrid.Location = new System.Drawing.Point(12, 157);
			this.dataGrid.Name = "dataGrid";
			this.dataGrid.ReadOnly = true;
			this.dataGrid.RowHeadersWidth = 100;
			this.dataGrid.RowTemplate.Height = 25;
			this.dataGrid.Size = new System.Drawing.Size(1061, 491);
			this.dataGrid.TabIndex = 0;
			// 
			// No
			// 
			this.No.Frozen = true;
			this.No.HeaderText = "No";
			this.No.Name = "No";
			this.No.ReadOnly = true;
			// 
			// FirstName
			// 
			this.FirstName.Frozen = true;
			this.FirstName.HeaderText = "First Name";
			this.FirstName.Name = "FirstName";
			this.FirstName.ReadOnly = true;
			// 
			// LastName
			// 
			this.LastName.HeaderText = "Last Name";
			this.LastName.Name = "LastName";
			this.LastName.ReadOnly = true;
			// 
			// Title
			// 
			this.Title.HeaderText = "Title";
			this.Title.Name = "Title";
			this.Title.ReadOnly = true;
			// 
			// CompanyName
			// 
			this.CompanyName.HeaderText = "Company Name";
			this.CompanyName.Name = "CompanyName";
			this.CompanyName.ReadOnly = true;
			// 
			// PersonLinkedinUrl
			// 
			this.PersonLinkedinUrl.HeaderText = "Person Linkedin Url";
			this.PersonLinkedinUrl.Name = "PersonLinkedinUrl";
			this.PersonLinkedinUrl.ReadOnly = true;
			// 
			// Email
			// 
			this.Email.HeaderText = "Email";
			this.Email.Name = "Email";
			this.Email.ReadOnly = true;
			// 
			// Website
			// 
			this.Website.HeaderText = "Website";
			this.Website.Name = "Website";
			this.Website.ReadOnly = true;
			// 
			// mobile_phone
			// 
			this.mobile_phone.HeaderText = "Mobile Phone";
			this.mobile_phone.Name = "mobile_phone";
			this.mobile_phone.ReadOnly = true;
			// 
			// btnLoad
			// 
			this.btnLoad.Location = new System.Drawing.Point(548, 38);
			this.btnLoad.Name = "btnLoad";
			this.btnLoad.Size = new System.Drawing.Size(75, 23);
			this.btnLoad.TabIndex = 1;
			this.btnLoad.Text = "Load";
			this.btnLoad.UseVisualStyleBackColor = true;
			this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
			// 
			// txtSpreadSheetId
			// 
			this.txtSpreadSheetId.Location = new System.Drawing.Point(107, 38);
			this.txtSpreadSheetId.Name = "txtSpreadSheetId";
			this.txtSpreadSheetId.Size = new System.Drawing.Size(435, 23);
			this.txtSpreadSheetId.TabIndex = 3;
			this.txtSpreadSheetId.Text = "1BPg9d6sAC6dv1oz1aplKKUWeQAw4q1xF7-cNAclMKNo";
			// 
			// txtSheetName
			// 
			this.txtSheetName.Location = new System.Drawing.Point(107, 67);
			this.txtSheetName.Name = "txtSheetName";
			this.txtSheetName.Size = new System.Drawing.Size(146, 23);
			this.txtSheetName.TabIndex = 5;
			this.txtSheetName.Text = "TMU - List 1";
			// 
			// txtColumnStart
			// 
			this.txtColumnStart.Location = new System.Drawing.Point(361, 67);
			this.txtColumnStart.Name = "txtColumnStart";
			this.txtColumnStart.Size = new System.Drawing.Size(23, 23);
			this.txtColumnStart.TabIndex = 7;
			this.txtColumnStart.Text = "A";
			// 
			// txtColumnEnd
			// 
			this.txtColumnEnd.Location = new System.Drawing.Point(408, 67);
			this.txtColumnEnd.Name = "txtColumnEnd";
			this.txtColumnEnd.Size = new System.Drawing.Size(22, 23);
			this.txtColumnEnd.TabIndex = 9;
			this.txtColumnEnd.Text = "G";
			// 
			// txtFirstRow
			// 
			this.txtFirstRow.Location = new System.Drawing.Point(516, 67);
			this.txtFirstRow.Name = "txtFirstRow";
			this.txtFirstRow.Size = new System.Drawing.Size(26, 23);
			this.txtFirstRow.TabIndex = 11;
			this.txtFirstRow.Text = "2";
			// 
			// btnPopulate
			// 
			this.btnPopulate.Location = new System.Drawing.Point(629, 38);
			this.btnPopulate.Name = "btnPopulate";
			this.btnPopulate.Size = new System.Drawing.Size(75, 23);
			this.btnPopulate.TabIndex = 12;
			this.btnPopulate.Text = "Populate";
			this.btnPopulate.UseVisualStyleBackColor = true;
			this.btnPopulate.Click += new System.EventHandler(this.btnPopulate_Click);
			// 
			// txtPDLKey
			// 
			this.txtPDLKey.Location = new System.Drawing.Point(107, 9);
			this.txtPDLKey.Name = "txtPDLKey";
			this.txtPDLKey.Size = new System.Drawing.Size(435, 23);
			this.txtPDLKey.TabIndex = 14;
			this.txtPDLKey.Text = "034f1ab06fa093e677368972975ac7ea0551ed372348d71c97a6943bc92802ec";
			// 
			// chkValidMobilePhone
			// 
			this.chkValidMobilePhone.AutoSize = true;
			this.chkValidMobilePhone.Location = new System.Drawing.Point(107, 102);
			this.chkValidMobilePhone.Name = "chkValidMobilePhone";
			this.chkValidMobilePhone.Size = new System.Drawing.Size(185, 19);
			this.chkValidMobilePhone.TabIndex = 15;
			this.chkValidMobilePhone.Text = "Valid \"Mobile Number\" record";
			this.chkValidMobilePhone.UseVisualStyleBackColor = true;
			// 
			// chkValidEmail
			// 
			this.chkValidEmail.AutoSize = true;
			this.chkValidEmail.Location = new System.Drawing.Point(107, 127);
			this.chkValidEmail.Name = "chkValidEmail";
			this.chkValidEmail.Size = new System.Drawing.Size(130, 19);
			this.chkValidEmail.TabIndex = 16;
			this.chkValidEmail.Text = "Valid \"Email\" record";
			this.chkValidEmail.UseVisualStyleBackColor = true;
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1085, 660);
			this.Controls.Add(this.chkValidEmail);
			this.Controls.Add(this.chkValidMobilePhone);
			this.Controls.Add(this.txtPDLKey);
			this.Controls.Add(label6);
			this.Controls.Add(this.btnPopulate);
			this.Controls.Add(this.txtFirstRow);
			this.Controls.Add(label5);
			this.Controls.Add(this.txtColumnEnd);
			this.Controls.Add(label4);
			this.Controls.Add(this.txtColumnStart);
			this.Controls.Add(label3);
			this.Controls.Add(this.txtSheetName);
			this.Controls.Add(label2);
			this.Controls.Add(this.txtSpreadSheetId);
			this.Controls.Add(label1);
			this.Controls.Add(this.btnLoad);
			this.Controls.Add(this.dataGrid);
			this.Name = "Form1";
			this.Text = "Tool";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGrid)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

		#endregion

		private DataGridView dataGrid;
		private Button btnLoad;
		private TextBox txtSpreadSheetId;
		private TextBox txtSheetName;
		private TextBox txtColumnStart;
		private TextBox txtColumnEnd;
		private TextBox txtFirstRow;
		private Button btnPopulate;
		private DataGridViewTextBoxColumn No;
		private DataGridViewTextBoxColumn FirstName;
		private DataGridViewTextBoxColumn LastName;
		private DataGridViewTextBoxColumn Title;
		private DataGridViewTextBoxColumn CompanyName;
		private DataGridViewTextBoxColumn PersonLinkedinUrl;
		private DataGridViewTextBoxColumn Email;
		private DataGridViewTextBoxColumn Website;
		private DataGridViewTextBoxColumn mobile_phone;
		private TextBox txtPDLKey;
		private CheckBox chkValidMobilePhone;
		private CheckBox chkValidEmail;
	}
}