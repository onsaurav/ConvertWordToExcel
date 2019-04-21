using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop;
using System.Reflection;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;

namespace ConvertWordToExcel
{
    public partial class frmConvertWordToExcel : Form
    {
        static object misValue = Missing.Value;

        public frmConvertWordToExcel()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void frmConvertWordToExcel_FormClosed(object sender, FormClosedEventArgs e)
        {
           System.Windows.Forms.Application.Exit();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            string ColText = "";
            string WordText = "";
            
            if (txtFrom.Text.Trim() == "")
            {
                MessageBox.Show("Sorry! Invalid source (word) file selected.", "File", MessageBoxButtons.OK);
                return;
            }
            if (txtTo.Text.Trim() == "")
            {
                MessageBox.Show("Sorry! Invalid destination (excel) file path selected.", "File", MessageBoxButtons.OK);
                return;
            }

            Workbook xlWB = new Workbook();
            
            Worksheet xlWorkSheet = new Worksheet();
            Application xlApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            
            xlWB=xlApp.Workbooks.Add(misValue);

            xlWorkSheet.Cells[0, 0] = "FIRST NAME";
            Microsoft.Office.Interop.Excel.Range range1 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[0,0];
            range1.EntireRow.Font.Bold = true;
            range1.EntireColumn.AutoFit();

            xlWorkSheet.Cells[0, 0] = "LAST NAME";
            Microsoft.Office.Interop.Excel.Range range2 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[0, 1];
            range2.EntireRow.Font.Bold = true;
            range2.EntireColumn.AutoFit();

            xlWorkSheet.Cells[0, 0] = "EMAIL";
            Microsoft.Office.Interop.Excel.Range range3 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[0, 2];
            range3.EntireRow.Font.Bold = true;
            range3.EntireColumn.AutoFit();

            xlWorkSheet.Cells[0, 0] = "PHONE";
            Microsoft.Office.Interop.Excel.Range range4 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[0, 3];
            range4.EntireRow.Font.Bold = true;
            range4.EntireColumn.AutoFit();

            xlWorkSheet.Cells[0, 0] = "EXPERIENCE";
            Microsoft.Office.Interop.Excel.Range range5 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[0, 4];
            range5.EntireRow.Font.Bold = true;
            range5.EntireColumn.AutoFit();

            xlWorkSheet.Cells[0, 0] = "CAPITAL";
            Microsoft.Office.Interop.Excel.Range range6 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[0, 5];
            range6.EntireRow.Font.Bold = true;
            range6.EntireColumn.AutoFit();

            bool AddState = false;
            int j = 0; int r = 0;
            WordText = AllText(txtFrom.Text.Trim());
            string[] stringSeparators = new string[] { "\r\n" };
            string[] Temp = WordText.Split(stringSeparators, StringSplitOptions.None);
            for (int i = 0; i < Temp.Length - 1; i++)
            {
                AddState = false;
                ColText = "";
                ColText = Temp[i].ToString();

                if (Temp[i].ToString().IndexOf("FIRST NAME") >= 0)
                {
                    j++; r++; AddState = true;
                    ColText.Replace("FIRST NAME", "");
                    ColText = ColText.Trim();
                }
                else if (Temp[i].ToString().IndexOf("LAST NAME") >= 0)
                {
                    r++; AddState = true;
                    ColText.Replace("LAST NAME", "");
                    ColText = ColText.Trim();
                }
                else if (Temp[i].ToString().IndexOf("EMAIL") >= 0)
                {
                    r++; AddState = true;
                    ColText.Replace("EMAIL", "");
                    ColText = ColText.Trim();
                }
                else if (Temp[i].ToString().IndexOf("PHONE") >= 0)
                {
                    r++; AddState = true;
                    ColText.Replace("PHONE", "");
                    ColText = ColText.Trim();
                }
                else if (Temp[i].ToString().IndexOf("EXPERIENCE") >= 0)
                {
                    r++; AddState = true;
                    ColText.Replace("EXPERIENCE", "");
                    ColText = ColText.Trim();
                }
                else if (Temp[i].ToString().IndexOf("CAPITAL") >= 0)
                {
                    r++; AddState = true;
                    ColText.Replace("CAPITAL", "");
                    ColText = ColText.Trim();
                }
                if (AddState == true)
                {
                    xlWorkSheet.Cells[r, j] = ColText.Trim();
                    Microsoft.Office.Interop.Excel.Range range7 = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[r, j];
                    range7.EntireRow.Font.Bold = false;
                    range7.EntireColumn.AutoFit();
                }
            }

            xlWB.SaveAs(txtTo.Text.Trim(), XlFileFormat.xlTextWindows , misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWB.Close(true, misValue, misValue);
            xlApp.Quit();
            ReleaseObject(xlWB);
            ReleaseObject(xlApp);   
        }

        private void btnConvertFrom_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog(); 
            dialog.Filter = "MSWord files (*.doc)|*.doc";
            dialog.InitialDirectory = "C:";
            dialog.Title = "Select a doc file";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtFrom.Text = dialog.FileName;
            }
            else
            {
                MessageBox.Show("Sorry! No file is seected.", "File", MessageBoxButtons.OK);
                txtFrom.Text = "";
            }           
        }

        private void btnConvertTo_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog oFBD = new FolderBrowserDialog();
            if (oFBD.ShowDialog() == DialogResult.OK)
            {
                txtTo.Text = FileNameTo(oFBD.SelectedPath,1);
            }
            else
            {
                MessageBox.Show("Sorry! No file psth is seected.", "File", MessageBoxButtons.OK);
                txtTo.Text = "";
            }
        }

        private string FileNameTo(string path, int i)
        {
            string GetExcelName = "";
            string FName = "Converted_Excel_" + i.ToString("00000")+ ".xls";
            if (path.Substring(path.Length - 1, 1) == "\\") path = path.Substring(0, path.Length - 1);
            FName = path + "\\" + FName;
            if (File.Exists(FName))
            {
                i++;
                GetExcelName = FileNameTo(path, i++);
            }
            else
            {
                GetExcelName = FName;
            }
            return GetExcelName;
        }

        private void frmConvertWordToExcel_Load(object sender, EventArgs e)
        {

        }

        private string AllText(string path)
        {
            string returnText = "";
            try
            {
                Microsoft.Office.Interop.Word.ApplicationClass wordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
                object file = path;
                object nullobj = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(ref file, ref nullobj, ref nullobj,
                                                                                    ref nullobj, ref nullobj, ref nullobj,
                                                                                    ref nullobj, ref nullobj, ref nullobj,
                                                                                    ref nullobj, ref nullobj, ref nullobj,
                                                                                    ref nullobj, ref nullobj, ref nullobj, ref nullobj);
                doc.ActiveWindow.Selection.WholeStory();
                doc.ActiveWindow.Selection.Copy();
                IDataObject data = Clipboard.GetDataObject();
                string allText = data.GetData(DataFormats.Text).ToString();
                doc.Close(ref nullobj, ref nullobj, ref nullobj);
                wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);
                returnText = allText;
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString(), "Read Doc", MessageBoxButtons.OK); }
            return returnText;
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
