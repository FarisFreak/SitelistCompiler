using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ClosedXML.Excel;
using static SLBR.TableParser;

namespace SLBR
{
    public class CRList : IEnumerable<CRList>
    {
        public string CRTech { get; set; }
        public string CRNumber { get; set; }
        public string CRTitle { get; set; }
        public string CRSiteId { get; set; }

        public string CRStatus { get; set; }

        public override string ToString()
        {
            return String.Format("CR Tech: {0} | CR Number: {1} | CR Title: {2} | CR SiteId: {3} | CR Status: {4}",CRTech , CRNumber, CRTitle, CRSiteId, CRStatus);
        }

        public IEnumerator<CRList> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }

    internal class Program
    {
        private static string GetNumbers(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }

        public static string ColumnAdress(int col)
        {
            if (col <= 26)
            {
                return Convert.ToChar(col + 64).ToString();
            }
            int div = col / 26;
            int mod = col % 26;
            if (mod == 0) { mod = 26; div--; }
            return ColumnAdress(div) + ColumnAdress(mod);
        }

        public static int ColumnNumber(string colAdress)
        {
            int[] digits = new int[colAdress.Length];
            for (int i = 0; i < colAdress.Length; ++i)
            {
                digits[i] = Convert.ToInt32(colAdress[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }

        [STAThread]
        static void Main(string[] args)
        {
            bool preventive = false;

            Console.Write("Is preventive? (y/n) ");

            if (Console.ReadKey().Key == ConsoleKey.Y)
                preventive = true;

            Console.WriteLine("");

            try
            {
                string[] _tech = { "3G", "4G" };
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Multiselect = true,
                    Title = "Select Sitelist file(s)",
                    Filter = "Excel Document|.xlsx;*.xlsx"
                };

                List<CRList> crList = new List<CRList>();

                using (dlg)
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        foreach (string filename in dlg.FileNames)
                        {
                            XLWorkbook wb = new XLWorkbook(filename);

                            using (wb)
                            {
                                var ws = wb.Worksheets.First();
                                var range = ws.RangeUsed().RowsUsed().Skip(1);


                                foreach (string tech in _tech)
                                {
                                    foreach (var row in range)
                                    {
                                        crList.Add(new CRList()
                                        {
                                            CRTech = tech,
                                            CRNumber = row.Cell(ColumnNumber("A")).GetString(),
                                            CRTitle = row.Cell(ColumnNumber("C")).GetString(),
                                            CRSiteId = row.Cell(ColumnNumber("B")).GetString(),
                                            CRStatus = row.Cell(ColumnNumber("O")).GetString()
                                        }) ;
                                    }
                                }
                            }
                        }
                    }
                    else return;
                }

                using (XLWorkbook newFile = new XLWorkbook())
                {
                    var worksheet = newFile.Worksheets.Add("Sheet1");

                    worksheet.Cell("A1").Value = "Tech";
                    worksheet.Cell("B1").Value = "CR Name/Activity";
                    worksheet.Cell("C1").Value = "Group";
                    worksheet.Cell("D1").Value = "SiteID";
                    if (preventive) worksheet.Cell("E1").Value = "Status";

                    string currentCR = "";
                    int currentCRCount = 0;
                    int appendCRNum = 0;

                    foreach (CRList cr in crList)
                    {
                        if (currentCR != cr.CRNumber)
                        {
                            currentCR = cr.CRNumber;
                            currentCRCount = 0;
                            appendCRNum = 0;
                        }

                        if (currentCRCount >= 500)
                        {
                            appendCRNum++;
                            currentCRCount = 0;
                        }

                        currentCRCount++;

                        int lastRow = worksheet.LastRowUsed().RowNumber();
                        int lastCol = 1;
                        worksheet.Cell(lastRow + 1, lastCol++).Value = cr.CRTech;
                        worksheet.Cell(lastRow + 1, lastCol++).Value = (appendCRNum == 0) ? cr.CRNumber : cr.CRNumber + "_" + appendCRNum.ToString("D2");

                        var forbiddenChars = new string[] { "'", "\"", "; ", ":", ".", ",", "#", "&", "(", ")", "[", "]" };
                        foreach (string fc in forbiddenChars)
                        {
                            cr.CRTitle = cr.CRTitle.Replace(fc, "");
                        }

                        worksheet.Cell(lastRow + 1, lastCol++).Value = cr.CRTitle;

                        cr.CRSiteId = GetNumbers(cr.CRSiteId);
                        //cr.CRSiteId = Regex.Replace(cr.CRSiteId, "[^a-zA-Z0-9 -]", "");
                        cr.CRSiteId = cr.CRSiteId.Replace(" ", string.Empty);

                        worksheet.Cell(lastRow + 1, lastCol++).Value = cr.CRSiteId;
                        if (preventive) worksheet.Cell(lastRow + 1, lastCol++).Value = cr.CRStatus;
                    }

                    var table = crList.ToStringTable(
                        new[] { "Tech", "CR Name/Activity", "Group", "SiteID", "Status" },
                        a => a.CRTech, a => a.CRNumber, a => a.CRTitle, a => a.CRSiteId, a => a.CRStatus
                    );

                    Console.WriteLine(table);

                    SaveFileDialog saveFileDialog = new SaveFileDialog
                    {
                        Title = "Save compiled Sitelist file",
                        Filter = "Excel Document|.xlsx;*.xlsx"
                    };

                    using (saveFileDialog)
                    {
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            newFile.SaveAs(@saveFileDialog.FileName);
                            Console.WriteLine("File successfuly saved at " + saveFileDialog.FileName);
                        }
                        else return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}