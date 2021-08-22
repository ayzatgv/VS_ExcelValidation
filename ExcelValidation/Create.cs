using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelValidation
{
    public partial class Create : Form
    {
        public List<string> excelFileData = new();

        public Create()
        {
            InitializeComponent();
        }

        private void Create_Load(object sender, EventArgs e)
        {
            excelFileData.Clear();
        }

        private void Button_Insert_Click(object sender, EventArgs e)
        {
            // input_1
            if (string.IsNullOrEmpty(ComboBox_1_1.SelectedItem.ToString()))
            {
                MessageBox.Show("Please fill out WO Type");
                return;
            }

            string input_1 = ComboBox_1_1.SelectedItem.ToString();

            // input_2
            if (string.IsNullOrEmpty(ComboBox_2_1.SelectedItem.ToString()))
            {
                MessageBox.Show("Please fill out Rack Type");
                return;
            }
            if (string.IsNullOrEmpty(TextBox_2_2.Text))
            {
                MessageBox.Show("Please fill out Rack Type");
                return;
            }

            string input_2 = ComboBox_2_1.SelectedItem.ToString() + " " + TextBox_2_2.Text.Trim() + " Rack";

            // input_3
            if (string.IsNullOrEmpty(TextBox_3_1.Text))
            {
                MessageBox.Show("Please fill out BB Brand");
                return;
            }

            string input_3 = TextBox_3_1.Text.Trim();

            // input_4
            if (string.IsNullOrEmpty(TextBox_4_2.Text))
            {
                MessageBox.Show("Please fill out BB Type");
                return;
            }

            string input_4 = "12V*" + TextBox_4_2.Text.Trim() + "AH";

            if (!Regex.Match(input_4, @"^([0-9]+V[*][0-9]+AH|N\/A)$").Success)
            {
                MessageBox.Show("BB Type must be 12V*150AH | N/A");
                return;
            }

            // input_5
            if (string.IsNullOrEmpty(TextBox_5_1.Text) || string.IsNullOrEmpty(TextBox_5_3.Text))
            {
                MessageBox.Show("Please fill out Quantity");
                return;
            }

            string input_5 = TextBox_5_1.Text.Trim() + "/" + TextBox_5_3.Text.Trim() + "PCs";

            if (!Regex.Match(input_5, @"^([0-9]+\/[0-9]+\s*(?i)PCs(?-i)|N\/A)$").Success)
            {
                MessageBox.Show("Quantity must be 8/8 PCs | N/A");
                return;
            }

            // input_6
            if (string.IsNullOrEmpty(ComboBox_1_1.SelectedItem.ToString()))
            {
                MessageBox.Show("Please fill out Support Time");
                return;
            }

            string input_6 = "(" + ComboBox_6_2.Text.Trim() + ")";

            if (!Regex.Match(input_6, @"^(([(][0]?[0-1][:]([0-5][0-9]|[6][0])[)])|[(]Van[)])$").Success)
            {
                MessageBox.Show("Support Time must be (0:10) < 2hour | N/A");
                return;
            }

            // input_7
            if (string.IsNullOrEmpty(TextBox_7_2.Text) || string.IsNullOrEmpty(TextBox_7_4.Text))
            {
                MessageBox.Show("Please fill out Desire");
                return;
            }

            string input_7 = "D:" + TextBox_7_2.Text.Trim() + "*" + TextBox_7_4.Text.Trim() + "AH";

            if (!Regex.Match(input_7, @"^D[:]([0-9]{1,}[*][0-9]{1,}AH|\s*N\/A)$").Success)
            {
                MessageBox.Show("Desire must be D:8*200AH | D:N/A");
                return;
            }

            // input_8
            if (string.IsNullOrEmpty(TextBox_8_2.Text))
            {
                MessageBox.Show("Please fill out Load");
                return;
            }

            string input_8 = "L:" + TextBox_8_2.Text.Trim() + "A";

            if (!Regex.Match(input_8, @"^L[:]([0-9]{1,}([.][0-9]{1,})?A|\s*N\/A)$").Success)
            {
                MessageBox.Show("Load must be L:89A | L:N/A");
                return;
            }

            // input_9
            if (string.IsNullOrEmpty(TextBox_9_2.Text))
            {
                MessageBox.Show("Please fill out Installation Date");
                return;
            }

            string input_9 = "I:" + TextBox_9_2.Text.Trim();

            if (!Regex.Match(input_9, @"I:([0-9]{8}|Before[0-9]{4}|\s*N\/A)$").Success)
            {
                MessageBox.Show("Installation Date must be I:20200507 | I:Before2018 | I:N/A");
                return;
            }

            // input_10
            if (string.IsNullOrEmpty(TextBox_10_2.Text))
            {
                MessageBox.Show("Please fill out Product Date");
                return;
            }

            string input_10 = "P:" + TextBox_10_2.Text.Trim();

            if (!Regex.Match(input_10, @"^P:([0-9]{8}|\s*N\/A)$").Success)
            {
                MessageBox.Show("Product Date must be P:20200507 | P:N/A");
                return;
            }

            // input_11
            if (string.IsNullOrEmpty(TextBox_11_2.Text) || string.IsNullOrEmpty(TextBox_11_4.Text))
            {
                MessageBox.Show("Please fill out PSU");
                return;
            }

            string input_11 = "PSU:";
            if (!string.IsNullOrEmpty(TextBox_11_10.Text) && !string.IsNullOrEmpty(TextBox_11_12.Text) && !string.IsNullOrEmpty(TextBox_11_6.Text) && !string.IsNullOrEmpty(TextBox_11_8.Text))
                input_11 += TextBox_11_2.Text.Trim() + "*" + TextBox_11_4.Text.Trim() + "+" + TextBox_11_6.Text.Trim() + "*" + TextBox_11_8.Text.Trim() + "+" + TextBox_11_10.Text.Trim() + "*" + TextBox_11_12.Text.Trim();
            else if (!string.IsNullOrEmpty(TextBox_11_6.Text) && !string.IsNullOrEmpty(TextBox_11_8.Text))
                input_11 += TextBox_11_2.Text.Trim() + "*" + TextBox_11_4.Text.Trim() + "+" + TextBox_11_6.Text.Trim() + "*" + TextBox_11_8.Text.Trim();
            else
                input_11 += TextBox_11_2.Text.Trim() + "*" + TextBox_11_4.Text.Trim();
            input_11 += "W";

            if (!Regex.Match(input_11, @"^PSU[:]([0-9]{1,}[*][0-9]{1,}W|[0-9]{1,}[*][0-9]{1,}W?[+][0-9]{1,}[*][0-9]{1,}W|[0-9]{1,}[*][0-9]{1,}W?[+][0-9]{1,}[*][0-9]{1,}W?[+][0-9]{1,}[*][0-9]{1,}W|[0]|\s*N\/A)$").Success)
            {
                MessageBox.Show("PSU must be PSU:4*3000W | PSU:0 | PSU:N/A");
                return;
            }

            excelFileData.Add(input_1 + " - " + input_2 + " - " + input_3 + " - " + input_4 + " - " + input_5 + " - " + input_6 + " - " + input_7 + " - " + input_8 + " - " + input_9 + " - " + input_10 + " - " + input_11);
            RichText.AppendText(input_1 + " - " + input_2 + " - " + input_3 + " - " + input_4 + " - " + input_5 + " - " + input_6 + " - " + input_7 + " - " + input_8 + " - " + input_9 + " - " + input_10 + " - " + input_11);
        }

        private void Button_Reset_Click(object sender, EventArgs e)
        {
            ComboBox_1_1.SelectedIndex = -1;
            ComboBox_2_1.SelectedIndex = -1;
            TextBox_2_2.Text = "";
            TextBox_3_1.Text = "";
            TextBox_4_2.Text = "";
            TextBox_5_1.Text = "";
            TextBox_5_3.Text = "";
            ComboBox_6_2.Text = "";
            TextBox_7_2.Text = "";
            TextBox_7_4.Text = "";
            TextBox_8_2.Text = "";
            TextBox_9_2.Text = "";
            TextBox_10_2.Text = "";
            TextBox_11_2.Text = "";
            TextBox_11_4.Text = "";
            TextBox_11_6.Text = "";
            TextBox_11_8.Text = "";
            TextBox_11_10.Text = "";
            TextBox_11_12.Text = "";
        }

        private void Button_Export_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.ActiveSheet;

                try
                {
                    for (int i = 0; i < excelFileData.Count; i++)
                    {
                        excelWorksheet.Cells[i + 1, 1] = excelFileData[i];
                    }

                    excelApp.ActiveWorkbook.SaveAs(@$"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\output.xls", Excel.XlFileFormat.xlWorkbookNormal);
                }
                catch (Exception)
                {
                    MessageBox.Show("Error");
                }
                finally
                {
                    excelWorkbook.Close();
                    excelApp.Quit();

                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            MessageBox.Show("Excel file created at " + @$"{Environment.GetFolderPath(Environment.SpecialFolder.Desktop)}\output.xls");
        }

        private void Create_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void Button_Clear_Click(object sender, EventArgs e)
        {
            RichText.Text = "";
            excelFileData.Clear();
        }
    }
}
