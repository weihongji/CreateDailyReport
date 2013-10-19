namespace CreateDailyReport
{
	partial class MainForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.components = new System.ComponentModel.Container();
			this.lblMonth = new System.Windows.Forms.Label();
			this.cmbMonth = new System.Windows.Forms.ComboBox();
			this.btnCreate = new System.Windows.Forms.Button();
			this.lblDir = new System.Windows.Forms.Label();
			this.txtDir = new System.Windows.Forms.TextBox();
			this.btnSelectDir = new System.Windows.Forms.Button();
			this.lblTemplate = new System.Windows.Forms.Label();
			this.btnSelectTemplate = new System.Windows.Forms.Button();
			this.txtTemplate = new System.Windows.Forms.TextBox();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnExit = new System.Windows.Forms.Button();
			this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbYear = new System.Windows.Forms.ComboBox();
			this.lblYear = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// lblMonth
			// 
			this.lblMonth.AutoSize = true;
			this.lblMonth.Location = new System.Drawing.Point(29, 129);
			this.lblMonth.Name = "lblMonth";
			this.lblMonth.Size = new System.Drawing.Size(37, 13);
			this.lblMonth.TabIndex = 6;
			this.lblMonth.Text = "&Month";
			// 
			// cmbMonth
			// 
			this.cmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbMonth.FormattingEnabled = true;
			this.cmbMonth.Items.AddRange(new object[] {
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December"});
			this.cmbMonth.Location = new System.Drawing.Point(84, 125);
			this.cmbMonth.Name = "cmbMonth";
			this.cmbMonth.Size = new System.Drawing.Size(85, 21);
			this.cmbMonth.TabIndex = 7;
			// 
			// btnCreate
			// 
			this.btnCreate.Location = new System.Drawing.Point(128, 186);
			this.btnCreate.Name = "btnCreate";
			this.btnCreate.Size = new System.Drawing.Size(75, 23);
			this.btnCreate.TabIndex = 10;
			this.btnCreate.Text = "&Create";
			this.btnCreate.UseVisualStyleBackColor = true;
			this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
			// 
			// lblDir
			// 
			this.lblDir.AutoSize = true;
			this.lblDir.Location = new System.Drawing.Point(29, 52);
			this.lblDir.Name = "lblDir";
			this.lblDir.Size = new System.Drawing.Size(49, 13);
			this.lblDir.TabIndex = 0;
			this.lblDir.Text = "&Directory";
			// 
			// txtDir
			// 
			this.txtDir.Location = new System.Drawing.Point(84, 48);
			this.txtDir.Name = "txtDir";
			this.txtDir.Size = new System.Drawing.Size(284, 20);
			this.txtDir.TabIndex = 1;
			this.toolTip1.SetToolTip(this.txtDir, "Directory where daily reports will be created in");
			// 
			// btnSelectDir
			// 
			this.btnSelectDir.Location = new System.Drawing.Point(386, 47);
			this.btnSelectDir.Name = "btnSelectDir";
			this.btnSelectDir.Size = new System.Drawing.Size(50, 23);
			this.btnSelectDir.TabIndex = 2;
			this.btnSelectDir.Text = "...";
			this.btnSelectDir.UseVisualStyleBackColor = true;
			this.btnSelectDir.Click += new System.EventHandler(this.btnSelectDir_Click);
			// 
			// lblTemplate
			// 
			this.lblTemplate.AutoSize = true;
			this.lblTemplate.Location = new System.Drawing.Point(29, 88);
			this.lblTemplate.Name = "lblTemplate";
			this.lblTemplate.Size = new System.Drawing.Size(51, 13);
			this.lblTemplate.TabIndex = 3;
			this.lblTemplate.Text = "&Template";
			// 
			// btnSelectTemplate
			// 
			this.btnSelectTemplate.Location = new System.Drawing.Point(386, 83);
			this.btnSelectTemplate.Name = "btnSelectTemplate";
			this.btnSelectTemplate.Size = new System.Drawing.Size(50, 23);
			this.btnSelectTemplate.TabIndex = 5;
			this.btnSelectTemplate.Text = "...";
			this.btnSelectTemplate.UseVisualStyleBackColor = true;
			this.btnSelectTemplate.Click += new System.EventHandler(this.btnSelectTemplate_Click);
			// 
			// txtTemplate
			// 
			this.txtTemplate.Location = new System.Drawing.Point(84, 84);
			this.txtTemplate.Name = "txtTemplate";
			this.txtTemplate.Size = new System.Drawing.Size(284, 20);
			this.txtTemplate.TabIndex = 4;
			this.toolTip1.SetToolTip(this.txtTemplate, "Path of daily report template file");
			// 
			// btnExit
			// 
			this.btnExit.Location = new System.Drawing.Point(265, 186);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(75, 23);
			this.btnExit.TabIndex = 11;
			this.btnExit.Text = "E&xit";
			this.btnExit.UseVisualStyleBackColor = true;
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.Filter = "Excel files(*.xls; *.xlsx)|*.xls;*.xlsx|All files(*.*)|*.*";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(119, 15);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(213, 13);
			this.label1.TabIndex = 12;
			this.label1.Text = "Create Daily Report Files of a Month";
			// 
			// cmbYear
			// 
			this.cmbYear.FormattingEnabled = true;
			this.cmbYear.Location = new System.Drawing.Point(245, 125);
			this.cmbYear.Name = "cmbYear";
			this.cmbYear.Size = new System.Drawing.Size(85, 21);
			this.cmbYear.TabIndex = 9;
			// 
			// lblYear
			// 
			this.lblYear.AutoSize = true;
			this.lblYear.Location = new System.Drawing.Point(206, 129);
			this.lblYear.Name = "lblYear";
			this.lblYear.Size = new System.Drawing.Size(29, 13);
			this.lblYear.TabIndex = 8;
			this.lblYear.Text = "&Year";
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(468, 233);
			this.Controls.Add(this.cmbYear);
			this.Controls.Add(this.lblYear);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnExit);
			this.Controls.Add(this.btnSelectTemplate);
			this.Controls.Add(this.txtTemplate);
			this.Controls.Add(this.lblTemplate);
			this.Controls.Add(this.btnSelectDir);
			this.Controls.Add(this.txtDir);
			this.Controls.Add(this.lblDir);
			this.Controls.Add(this.btnCreate);
			this.Controls.Add(this.cmbMonth);
			this.Controls.Add(this.lblMonth);
			this.Name = "MainForm";
			this.Text = "Create Daily Reports";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label lblMonth;
		private System.Windows.Forms.ComboBox cmbMonth;
		private System.Windows.Forms.Button btnCreate;
		private System.Windows.Forms.Label lblDir;
		private System.Windows.Forms.TextBox txtDir;
		private System.Windows.Forms.Button btnSelectDir;
		private System.Windows.Forms.Label lblTemplate;
		private System.Windows.Forms.Button btnSelectTemplate;
		private System.Windows.Forms.TextBox txtTemplate;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cmbYear;
		private System.Windows.Forms.Label lblYear;
	}
}

