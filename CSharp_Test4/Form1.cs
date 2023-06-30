using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Text;

namespace CSharp_Test4
{
    public partial class Form1 : Form
    {
        private Label accountLabel;
        private Label passwordLabel;
        private TextBox accountTextBox;
        private TextBox passwordTextBox;
        private Button loginButton;
        private Label inputAccount;
        private Label inputOldPassword;
        private Label inputNewPassword;
        private Label confirmNewPassword;
        private TabControl UserTabControl;
        private TextBox inputAccountTextBox;
        private TextBox inputOldPasswordTextBox;
        private TextBox inputNewPasswordTextBox;
        private TextBox confirmNewPasswordTextBox;
        private Button changePasswordButton;
        private Button changePassword2Button;
        private ToolStrip ExcelToolStrip;
        private Panel bottomPanel;
        private TextBox xmlTextBox;
        private Dictionary<string, string> accountPassword;
        ToolStripButton choiceXmlButton = new ToolStripButton();
        ToolStripButton exportXmlButton = new ToolStripButton();
        Form newForm = new Form();
        private string selectedFilePath;


        public Form1()
        {
            accountPassword = new Dictionary<string, string>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            this.Size = new Size(900, 600);

            UserTabControl = new TabControl();
            UserTabControl.Multiline = false;
            UserTabControl.Appearance = TabAppearance.Normal;
            TabPage tabPage = new TabPage();
            tabPage.Text = "使用者管理";
            tabPage.BackColor = Color.LightBlue;
            UserTabControl.TabPages.Add(tabPage);
            this.Controls.Add(UserTabControl);

            changePasswordButton = new Button();
            changePasswordButton.Text = "更改密碼";
            changePasswordButton.Size = new Size(100, 40);
            changePasswordButton.Location = new Point(40, 20);
            changePasswordButton.BackColor = SystemColors.Control;
            changePasswordButton.Click += changePassword_Click;
            tabPage.Controls.Add(changePasswordButton);

            accountLabel = new Label();
            accountLabel.Text = "帳號 : ";
            accountLabel.AutoSize = true;
            accountLabel.Location = new Point(50, 130);
            accountLabel.Font = new Font(accountLabel.Font.FontFamily, 14, FontStyle.Regular);
            this.Controls.Add(accountLabel);

            accountTextBox = new TextBox();
            accountTextBox.Size = new Size(150, 60);
            accountTextBox.Location = new Point(113, 130);
            this.Controls.Add(accountTextBox);

            passwordLabel = new Label();
            passwordLabel.Text = "密碼 : ";
            passwordLabel.AutoSize = true;
            passwordLabel.Location = new Point(50, 180);
            passwordLabel.Font = new Font(passwordLabel.Font.FontFamily, 14, FontStyle.Regular);
            this.Controls.Add(passwordLabel);

            passwordTextBox = new TextBox();
            passwordTextBox.Size = new Size(150, 60);
            passwordTextBox.Location = new Point(113, 180);
            this.Controls.Add(passwordTextBox);

            loginButton = new Button();
            loginButton.Text = "登入";
            loginButton.Size = new Size(70, 50);
            loginButton.Location = new Point(300, 150);
            loginButton.Click += loginButton_Click;
            this.Controls.Add(loginButton);

            bottomPanel = new Panel();
            bottomPanel.Size = new Size(600, 300);
            bottomPanel.Dock = DockStyle.Bottom;
            bottomPanel.BackColor = Color.Gray;
            this.Controls.Add(bottomPanel);

            xmlTextBox = new TextBox();
            xmlTextBox.Multiline = true;
            xmlTextBox.Font = new Font(xmlTextBox.Font.FontFamily, 14, FontStyle.Regular);
            xmlTextBox.Size = new Size(600, 273);
            xmlTextBox.BackColor = Color.Gray;
            xmlTextBox.BorderStyle = BorderStyle.None;
            xmlTextBox.Dock = DockStyle.Top;
            xmlTextBox.Text = "灰色區域";
            bottomPanel.Controls.Add(xmlTextBox);

            ExcelToolStrip = new ToolStrip();

            Bitmap image = SystemIcons.Shield.ToBitmap();
            choiceXmlButton.Image = image;
            exportXmlButton.Image = image;
            choiceXmlButton.Visible = false;
            exportXmlButton.Visible = false;
            ExcelToolStrip.Size = new Size(600, 40);
            ExcelToolStrip.Location = new Point(0, 300);
            exportXmlButton.Dock = DockStyle.Top;
            choiceXmlButton.Click += choiceXmlButton_Click;
            exportXmlButton.Click += exportXmlButton_Click;
            ExcelToolStrip.Items.Add(choiceXmlButton);
            ExcelToolStrip.Items.Add(exportXmlButton);
            bottomPanel.Controls.Add(ExcelToolStrip);

        }

        private void exportXmlButton_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel Files (*.xlsx)|*.xlsx";
            save.Title = "儲存檔案";
            save.FileName = Path.GetFileName(selectedFilePath);

            if(save.ShowDialog() == DialogResult.OK)
            {
                string saveFilePath = save.FileName;
                using(var newWordboox = new OfficeOpenXml.ExcelPackage())
                {
                    var newWorksheet = newWordboox.Workbook.Worksheets.Add("sheet1");
                    string[] lines = xmlTextBox.Text.Split(new string[]
                    {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);

                    for(int i = 0; i < lines.Length; i++)
                    {
                        string[] values = lines[i].Split('\t');
                        for (int j = 0; j < values.Length; j++)
                        {
                            newWorksheet.Cells[i + 1, j + 1].Value = values[j];
                        }
                    }
                    newWordboox.SaveAs(new FileInfo(saveFilePath));

                }
                MessageBox.Show("檔案儲存成功");
            }
        }

        private void loginButton_Click(object sender, EventArgs e)
        {
            string account = accountTextBox.Text;
            string password = passwordTextBox.Text;
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFilePath = Path.Combine(desktopPath,"使用者帳密.xlsx");
            
            if (File.Exists(excelFilePath))
            {
                using (ExcelPackage package2 = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet2 = package2.Workbook.Worksheets[0];
                    int rowCount2 = worksheet2.Dimension.Rows;
                    int colCount2 = worksheet2.Dimension.Columns;

                    bool accountFound = false;
                    bool passwordMatched = false;

                    for(int row2 = 2; row2 <= rowCount2; row2++)
                    {
                        string accountFormSheet = worksheet2.Cells[row2, 1].Value?.ToString();
                        string passwordFormSheet = worksheet2.Cells[row2, 2].Value?.ToString();

                        if(accountFormSheet == account)
                        {
                            accountFound = true;

                            if (passwordFormSheet == password)
                            {
                                passwordMatched = true;
                                break;
                            }
                        }
                    }
                    if(accountFound && passwordMatched)
                    {
                        MessageBox.Show("登入成功!");
                        choiceXmlButton.Visible = true;
                        exportXmlButton.Visible = true;
                    }
                    else
                    {
                        MessageBox.Show("帳號或密碼錯誤!");
                    }
                }
            }
            else
            {
                MessageBox.Show("使用者帳密.xlsx檔案不存在");
            }
        }

        private void OpenFileDialog()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "Excel Files(*.xlsx)|*.xlsx";

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFile.FileName;
                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        if (worksheet != null)
                        {
                            int rowCount = worksheet.Dimension.Rows;
                            int colCount = worksheet.Dimension.Columns;
                            StringBuilder contentBulider = new StringBuilder();
                            for (int row = 1; row <= rowCount; row++)
                            {
                                for (int col = 1; col <= colCount; col++)
                                {
                                    object cellValue = worksheet.Cells[row, col].Value;
                                    contentBulider.Append(cellValue?.ToString() ?? "");
                                    contentBulider.Append("\t");
                                }
                                contentBulider.AppendLine();
                            }
                            xmlTextBox.Text = contentBulider.ToString();
                        }
                        else
                        {
                            MessageBox.Show("指定的工作表不存在!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("無法讀取檔案 : " + ex.Message);
                }
            }
        }

        private void choiceXmlButton_Click(object sneder, EventArgs e)
        {
            OpenFileDialog();
        }

        private void changePassword_Click(object sender , EventArgs e)
        {
            newForm.Visible = false;
            newForm.Text = "更改密碼";
            newForm.Size = new Size(400, 500);

            inputAccount = new Label();
            inputAccount.Text = "輸入帳號 : ";
            inputAccount.AutoSize = true;
            inputAccount.Location = new Point(40, 70);
            newForm.Controls.Add(inputAccount);

            inputAccountTextBox = new TextBox();
            inputAccountTextBox.Size = new Size(180, 40);
            inputAccountTextBox.Location = new Point(160, 70);
            newForm.Controls.Add(inputAccountTextBox);

            inputOldPassword = new Label();
            inputOldPassword.Text = "輸入舊密碼 : ";
            inputOldPassword.AutoSize = true;
            inputOldPassword.Location = new Point(40, 130);
            newForm.Controls.Add(inputOldPassword);

            inputOldPasswordTextBox = new TextBox();
            inputOldPasswordTextBox.Size = new Size(180, 40);
            inputOldPasswordTextBox.Location = new Point(160, 130);
            newForm.Controls.Add(inputOldPasswordTextBox);

            inputNewPassword = new Label();
            inputNewPassword.Text = "輸入新密碼 : ";
            inputNewPassword.AutoSize = true;
            inputNewPassword.Location = new Point(40, 190);
            newForm.Controls.Add(inputNewPassword);

            inputNewPasswordTextBox = new TextBox();
            inputNewPasswordTextBox.Size = new Size(180, 40);
            inputNewPasswordTextBox.Location = new Point(160, 190);
            newForm.Controls.Add(inputNewPasswordTextBox);

            confirmNewPassword = new Label();
            confirmNewPassword.Text = "確認新密碼 : ";
            confirmNewPassword.AutoSize = true;
            confirmNewPassword.Location = new Point(40, 250);
            newForm.Controls.Add(confirmNewPassword);

            confirmNewPasswordTextBox = new TextBox();
            confirmNewPasswordTextBox.Size = new Size(180, 40);
            confirmNewPasswordTextBox.Location = new Point(160, 250);
            newForm.Controls.Add(confirmNewPasswordTextBox);

            changePassword2Button = new Button();
            changePassword2Button.Size = new Size(120, 50);
            changePassword2Button.Location = new Point(130, 330);
            changePassword2Button.Text = "確認更改";
            changePassword2Button.Click += changePasswordButton2_Click;
            newForm.Controls.Add(changePassword2Button);

        newForm.Show();
        }

        private void changePasswordButton2_Click(object sender,EventArgs e)
        {
            string account = inputAccountTextBox.Text;
            string OldPassword = inputOldPasswordTextBox.Text;
            string newPassword = inputNewPasswordTextBox.Text;
            string confirmNewPassword = confirmNewPasswordTextBox.Text;

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "使用者帳密.xlsx");

            if (File.Exists(filePath))
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    bool accountFound = false;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string accountFromSheet = worksheet.Cells[row, 1].Value?.ToString();
                        string passwordFromSheet = worksheet.Cells[row, 2].Value?.ToString();

                        if (accountFromSheet == account)
                        {
                            accountFound = true;

                            if (passwordFromSheet == OldPassword)
                            {
                                if (newPassword == confirmNewPassword)
                                {
                                    // 修改密碼
                                    worksheet.Cells[row, 2].Value = newPassword.ToString();
                                    package.Save();

                                    MessageBox.Show("密碼修改成功！");
                                }
                                else
                                {
                                    MessageBox.Show("新密碼與確認新密碼不相符！");
                                }

                                break;
                            }
                            else
                            {
                                MessageBox.Show("目前密碼不正確！");
                            }
                        }
                    }

                    if (!accountFound)
                    {
                        MessageBox.Show("帳號不存在！");
                    }
                }
                newForm.Close();
            }
            else
            {
                MessageBox.Show("帳號密碼檔案遺失");
                newForm.Close();
            }
        }

    }
}
