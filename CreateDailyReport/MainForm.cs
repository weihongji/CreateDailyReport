using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace CreateDailyReport
{
	public partial class MainForm : Form
	{
		string m_FiledDirectory = "";
		string m_FiledTemplatePath = "";

		public MainForm() {
			InitializeComponent();
		}

		private string FiledDir {
			get {
				if (m_FiledDirectory == "") {
					if (File.Exists("dr.txt")) {
						string[] lines = File.ReadAllLines("dr.txt");
						if (lines.Length > 0) {
							m_FiledDirectory = lines[0].Trim();
						}
					}
					if (m_FiledDirectory.Length==0) {
						m_FiledDirectory = Directory.GetCurrentDirectory();
						File.WriteAllText("dr.txt", m_FiledDirectory);
					}
				}
				return m_FiledDirectory;
			}
			set {
				if (m_FiledDirectory != txtDir.Text) {
					string s = txtDir.Text;
					string[] lines = File.ReadAllLines("dr.txt");
					if (lines.Length > 1) {
						s += "\r\n" + lines[1].Trim();
					}
					File.WriteAllText("dr.txt", s);
					m_FiledDirectory = txtDir.Text;
				}
			}
		}

		private string FiledTemplatePath {
			get {
				if (m_FiledTemplatePath == "") {
					if (File.Exists("dr.txt")) {
						string[] lines;
						lines = File.ReadAllLines("dr.txt");
						if (lines.Length > 1) {
							m_FiledTemplatePath = lines[1].Trim();
						}
					}
				}
				return m_FiledTemplatePath;
			}
			set {
				if (m_FiledTemplatePath != txtTemplate.Text) {
					string s = File.ReadAllLines("dr.txt")[0].Trim() + "\r\n" + txtTemplate.Text;
					File.WriteAllText("dr.txt", s);
					m_FiledTemplatePath = txtTemplate.Text;
				}
			}
		}

		private void MainForm_Load(object sender, EventArgs e) {
			txtDir.Text = FiledDir;
			txtTemplate.Text = FiledTemplatePath;
			// Initial Year combobox
			for (int i = 0; i < 5; i++) {
				cmbYear.Items.Add(DateTime.Today.Year + i);
			}
			cmbYear.Text = DateTime.Today.Year.ToString();
			cmbMonth.SelectedIndex = DateTime.Today.Month - 1;
		}

		private void btnSelectDir_Click(object sender, EventArgs e) {
			folderBrowserDialog1.SelectedPath = txtDir.Text;
			if (folderBrowserDialog1.ShowDialog() == DialogResult.OK) {
				txtDir.Text = folderBrowserDialog1.SelectedPath;
			}
		}

		private void btnSelectTemplate_Click(object sender, EventArgs e) {
			string currentDir = Directory.GetCurrentDirectory();
			string initTemplateDir;

			if (txtTemplate.Text.Length == 0 || !File.Exists(txtTemplate.Text)) {
				initTemplateDir = txtDir.Text;
			}
			else {
				FileInfo fInfo = new FileInfo(txtTemplate.Text);
				initTemplateDir = fInfo.DirectoryName;
			}
			openFileDialog1.InitialDirectory = initTemplateDir;
			if (openFileDialog1.ShowDialog() == DialogResult.OK) {
				txtTemplate.Text = openFileDialog1.FileName;
			}
			Directory.SetCurrentDirectory(currentDir);
			
		}

		private void btnCreate_Click(object sender, EventArgs e) {
			if (CreateDR()) {
				FiledDir = txtDir.Text;
				FiledTemplatePath = txtTemplate.Text;
			}
		}

		private void btnExit_Click(object sender, EventArgs e) {
			System.Windows.Forms.Application.Exit();
		}

		private bool CreateDR(bool includeWeekend) {
			if (!Directory.Exists(txtDir.Text)) {
				MessageBox.Show("Destination directory does not exist.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
				txtDir.SelectAll();
				txtDir.Focus();
				return false;
			}
			if (!File.Exists(txtTemplate.Text)) {
				MessageBox.Show("Daily report template file does not exist.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
				txtTemplate.SelectAll();
				txtTemplate.Focus();
				return false;
			}
			bool validYear = false;
			if (IsDigits(cmbYear.Text)) {
				int iYear = Convert.ToInt16(cmbYear.Text);
				if (iYear >= 1000 && iYear <= 9999) {
					validYear = true;
				}
			}
			if (!validYear) {
				MessageBox.Show("The year entered is not valid.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
				cmbYear.SelectAll();
				cmbYear.Focus();
				return false;
			}
			// Create the month folder
			string monthFolder = txtDir.Text;
			if (monthFolder.Substring(monthFolder.Length - 1) == "\\") {
				monthFolder += cmbMonth.Text;
			}
			else {
				monthFolder += "\\" + cmbMonth.Text;
			}
			if (!Directory.Exists(monthFolder)) {
				Directory.CreateDirectory(monthFolder);
			}

			// Create daily report files of the month
			int month = cmbMonth.SelectedIndex + 1;
			int year = Convert.ToInt16(cmbYear.Text);
			string newFileName;
			string fileExtension=(new FileInfo(txtTemplate.Text)).Extension;
			DateTime dt = new DateTime(year, month, 1);
			int createdCount = 0, jumpedCount = 0;
			Microsoft.Office.Interop.Excel.Application theExcelApp = new Microsoft.Office.Interop.Excel.Application();
			Workbook theExcelBook;
			Worksheet theSheet;
			Range theCell;

			do {
				if (!includeWeekend) {
					if (dt.DayOfWeek == DayOfWeek.Saturday || dt.DayOfWeek == DayOfWeek.Sunday) {
						dt = dt.AddDays(1);
						continue;
					}
				}
				newFileName = monthFolder + "\\" + dt.ToString("yyyyMMdd")+fileExtension;
				if (!File.Exists(newFileName)) {
					// Create file
					File.Copy(txtTemplate.Text, newFileName);
					createdCount++;
					// Update Report Date
					try {
						theExcelBook = theExcelApp.Workbooks.Open(newFileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
						theSheet = (Worksheet)theExcelBook.Sheets.get_Item("Sheet1");
						theCell = (Range)theSheet.get_Range("D2", "D2");

						theCell.Value2 = dt.ToString("MM/dd/yyyy");

						theExcelBook.Save();
						theExcelApp.Workbooks.Close();
					}
					catch (Exception e) {
						MessageBox.Show(e.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				else {
					jumpedCount++;
				}

				dt = dt.AddDays(1);
			} while (dt.Month == month && dt.Year == year);

			// Close excel application
			theExcelApp.Quit();

			string result;
			if (jumpedCount == 0) {
				if (createdCount <= 1) {
					result = createdCount.ToString() + " daily report file has been created.";
				}
				else {
					result = createdCount.ToString() + " daily report files have been created.";
				}
			}
			else {
				result = createdCount.ToString() + " daily report file has been created, and " + jumpedCount.ToString() + " jumped for existence.";
			}
			MessageBox.Show(result, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

			return true;
		}

		private bool CreateDR() {
			return CreateDR(false);
		}

		private bool IsDigits(string s) {
			for (int i = 0; i < s.Length; i++) {
				if (!char.IsDigit(s[i])) {
					return false;
				}
			}
			if (s.Length > 0) {
				return true;
			}
			else {
				return false;
			}
		}
	}
}