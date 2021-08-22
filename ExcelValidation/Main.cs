using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelValidation
{
    public partial class Main : Form
    {
        public string excelFilePath;

        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
        }

        private void Button_Import_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new()
            {
                InitialDirectory = @"C:\",
                Title = "Select Excel File",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "xlsx",
                Filter = "CSV Files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = openFileDialog.FileName;
            }
        }

        private void Button_Validate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("import excel file first");
            }
            else
            {
                Excel.Application xlApp = new();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excelFilePath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                try
                {
                    for (int i = 2; i <= xlRange.Rows.Count; i++)
                    {
                        string text = "";
                        string site = "";

                        if (xlRange.Cells[i, 2].Value2 != null)
                        {
                            text = xlRange.Cells[i, 2].Value2.ToString();
                        }
                        if (xlRange.Cells[i, 1].Value2 !=null)
                        {
                            site = xlRange.Cells[i, 1].Value2.ToString();
                        }

                        if (string.IsNullOrEmpty(text))
                        {
                            WriteRichText(i, site, text, "Battery Vandalism - GSM Eltek Rack - Narada - 12V*155AH - 4/4 PCs - (Van) - D:8*200AH - L:57A - I:20150429 - P:N/A - PSU:3*2000W");
                            continue;
                        }

                        if (text.Contains("HWS"))
                        {
                            int index = text.IndexOf("-");
                            text = (index < 0) ? text : text.Remove(0, index + 1);
                        }

                        string[] subText = text.Split('-');
                        if(subText.Length < 11)
                        {
                            WriteRichText(i, site, text, "Battery Vandalism - GSM Eltek Rack - Narada - 12V*155AH - 4/4 PCs - (Van) - D:8*200AH - L:57A - I:20150429 - P:N/A - PSU:3*2000W");
                            continue;
                        }

                        // [0]
                        if (!subText[0].Trim().Equals("Battery Installation") && !subText[0].Trim().Equals("Battery Vandalism") && !subText[0].Trim().Equals("Battery Second Hand") && !subText[0].Trim().Equals("Battery Reshuffle"))
                        {
                            WriteRichText(i, site, subText[0], "Battery Installation | Battery Vandalism | Battery Second Hand");
                        }
                        // [1]
                        if (string.IsNullOrEmpty(subText[1]))
                        {
                            WriteRichText(i, site, subText[1], "GSM Minishelter Rack | N/A");
                        }
                        // [2]
                        if (string.IsNullOrEmpty(subText[2]))
                        {
                            WriteRichText(i, site, subText[2], "Dengta | N/A");
                        }
                        // [3]
                        if (string.IsNullOrEmpty(subText[3]) || !Regex.Match(subText[3].Trim(), @"^([0-9]+V[*][0-9]+AH|N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[3], "12V*150AH | N/A");
                        }
                        // [4]
                        if (string.IsNullOrEmpty(subText[4]) || !Regex.Match(subText[4].Trim(), @"^([0-9]+\/[0-9]+\s*(?i)PCs(?-i)|N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[4], "8/8 PCs | N/A");
                        }
                        // [5]
                        if (!subText[5].Equals("VAN") && !Regex.Match(subText[5].Trim(), @"^(([(][0]?[0-1][:]([0-5][0-9]|[6][0])[)])|[(]Van[)])$").Success)
                        {
                            WriteRichText(i, site, subText[5], "(0:10) < 2hour | N/A");
                        }
                        // [6]
                        if (string.IsNullOrEmpty(subText[6]) || !Regex.Match(subText[6].Trim(), @"^D[:]([0-9]{1,}[*][0-9]{1,}AH|\s*N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[6], "D:8*200AH | D:N/A");
                        }
                        // [7]
                        if (string.IsNullOrEmpty(subText[7]) || !Regex.Match(subText[7].Trim(), @"^L[:]([0-9]{1,}([.][0-9]{1,})?A|\s*N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[7], "L:89A | L:N/A");
                        }
                        // [8]
                        if (string.IsNullOrEmpty(subText[8]) || !Regex.Match(subText[8].Trim(), @"I:([0-9]{8}|Before[0-9]{4}|\s*N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[8], "I:20200507 | I:Before2018 | I:N/A");
                        }
                        // [9]
                        if (string.IsNullOrEmpty(subText[9]) || !Regex.Match(subText[9].Trim(), @"^P:([0-9]{8}|\s*N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[9], "P:20200507 | P:N/A");
                        }
                        // [10]
                        if (string.IsNullOrEmpty(subText[10]) || !Regex.Match(subText[10].Trim(), @"^PSU[:]([0-9]{1,}[*][0-9]{1,}W|[0-9]{1,}[*][0-9]{1,}W?[+][0-9]{1,}[*][0-9]{1,}W|[0-9]{1,}[*][0-9]{1,}W?[+][0-9]{1,}[*][0-9]{1,}W?[+][0-9]{1,}[*][0-9]{1,}W|[0]|\s*N\/A)$").Success)
                        {
                            WriteRichText(i, site, subText[10], "PSU:4*3000W | PSU:0 | PSU:N/A");
                        }
                        Label_ErrorCount.Text = RichText.Lines.Length.ToString();
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error");
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);

                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);

                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);

                    MessageBox.Show("OK");
                }
            }
        }

        private void Button_Create_Click(object sender, EventArgs e)
        {
            var form = new Create
            {
                Location = this.Location,
                StartPosition = FormStartPosition.Manual
            };
            form.FormClosing += delegate { this.Show(); };
            form.Show();
            this.Hide();
        }

        private void Button_Clear_Click(object sender, EventArgs e)
        {
            excelFilePath = "";
            RichText.Text = "";
            Label_ErrorCount.Text = "0";
        }

        public void WriteRichText(int errorRow, string site, string errorMessage, string correctMessage)
        {
            if (!string.IsNullOrWhiteSpace(RichText.Text))
            {
                RichText.AppendText("\r\n" + "Row: " + errorRow + "| Site ID: " + site + "| Error: " + errorMessage + "| Correct Format: " + correctMessage);
            }
            else
            {
                RichText.AppendText("Row: " + errorRow + "| Site ID: " + site + "| Error: " + errorMessage + "| Correct Format: " + correctMessage);
            }
            RichText.ScrollToCaret();
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
    }
}
