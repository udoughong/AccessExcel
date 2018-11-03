using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp4AccessExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //範例一，簡單產生Excel檔案的方法
        private void CreateExcelFile(string excelVer)
        {
            IWorkbook wb = null;
            ISheet ws = null;

            switch (excelVer)
            {
                case ("2003"):
                    //建立Excel 2003檔案
                    wb = new HSSFWorkbook();
                    break;
                case ("2007"):
                    //建立excel 2007檔案
                    wb = new XSSFWorkbook();
                    break;
                default:
                    break;
            }

            ws = wb.CreateSheet("Sheet1");
            ws.CreateRow(0);//第一行為欄位名稱
            ws.GetRow(0).CreateCell(0).SetCellValue("name");
            ws.GetRow(0).CreateCell(1).SetCellValue("score");
            ws.CreateRow(1);//第二行之後為資料
            ws.GetRow(1).CreateCell(0).SetCellValue("abey");
            ws.GetRow(1).CreateCell(1).SetCellValue(85);
            ws.CreateRow(2);
            ws.GetRow(2).CreateCell(0).SetCellValue("tina");
            ws.GetRow(2).CreateCell(1).SetCellValue(82);
            ws.CreateRow(3);
            ws.GetRow(3).CreateCell(0).SetCellValue("boi");
            ws.GetRow(3).CreateCell(1).SetCellValue(84);
            ws.CreateRow(4);
            ws.GetRow(4).CreateCell(0).SetCellValue("hebe");
            ws.GetRow(4).CreateCell(1).SetCellValue(86);
            ws.CreateRow(5);
            ws.GetRow(5).CreateCell(0).SetCellValue("paul");
            ws.GetRow(5).CreateCell(1).SetCellValue(82);

            FileStream file = null;

            switch (excelVer)
            {
                case ("2003"):
                    file = new FileStream(@"c:\tmp\npoi.xls", FileMode.Create);//產生Excel 2003檔案
                    break;
                case ("2007"):
                    file = new FileStream(@"c:\tmp\npoi.xlsx", FileMode.Create);//產生Excel 2007檔案
                    break;
                default:
                    break;
            }

            wb.Write(file);
            file.Close();
        }

        //範例二，DataTable轉成Excel檔案的方法
        private void DataTableToExcelFile(DataTable dt, string excelVer)
        {
            IWorkbook wb = null;
            ISheet ws = null;

            switch (excelVer)
            {
                case ("2003"):
                    //建立Excel 2003檔案
                    wb = new HSSFWorkbook();
                    break;
                case ("2007"):
                    //建立excel 2007檔案
                    wb = new XSSFWorkbook();
                    break;
                default:
                    break;
            }

            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            FileStream file = null;

            switch (excelVer)
            {
                case ("2003"):
                    file = new FileStream(@"c:\tmp\npoi.xls", FileMode.Create);//產生Excel 2003檔案
                    break;
                case ("2007"):
                    file = new FileStream(@"c:\tmp\npoi.xlsx", FileMode.Create);//產生Excel 2007檔案
                    break;
                default:
                    break;
            }

            wb.Write(file);
            file.Close();
        }

        private void BtnCreate2k3Excel_Click(object sender, EventArgs e)
        {
            CreateExcelFile("2003");
        }

        private void BtnCreate2k7Excel_Click(object sender, EventArgs e)
        {
            CreateExcelFile("2007");
        }

        //補充範例、取得檔名的方法
        private void GetFileName(out string newFileName)
        {
            OpenFileDialog sflg = new OpenFileDialog();
            sflg.Filter = "Excel(*.xls)|*.xls|Excel(*.xlsx)|*.xlsx";
            if (sflg.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            {
                newFileName = "";
            }
            newFileName = sflg.FileName;
        }

        //範例三，Excel檔案轉成DataTable的方法
        private DataTable GetDataTableFromExcelFile(string fileName)
        {
            FileStream fs = null;
            DataTable dt = new DataTable();
            try
            {
                IWorkbook wb = null;
                fs = File.Open(fileName, FileMode.Open, FileAccess.Read);
                switch (Path.GetExtension(fileName).ToUpper())
                {
                    case ".XLS":
                        wb = new HSSFWorkbook(fs);
                        break;
                    case ".XLSX":
                        wb = new XSSFWorkbook(fs);
                        break;
                }
                if (wb.NumberOfSheets > 0)
                {
                    ISheet sheet = wb.GetSheetAt(0);
                    IRow headerRow = sheet.GetRow(0);

                    //處理標題列
                    for (int i = headerRow.FirstCellNum; i < headerRow.LastCellNum; i++)
                    {
                        dt.Columns.Add(headerRow.GetCell(i).StringCellValue.Trim());
                    }
                    IRow row = null;
                    DataRow dr = null;
                    CellType ct = CellType.Blank;
                    //標題列之後的資料
                    for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                    {
                        dr = dt.NewRow();
                        row = sheet.GetRow(i);
                        if (row == null) continue;
                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            ct = row.GetCell(j).CellType;
                            //如果此欄位格式為公式 則去取得CachedFormulaResultType
                            if (ct == CellType.Formula)
                            {
                                ct = row.GetCell(j).CachedFormulaResultType;
                            }
                            if (ct == CellType.Numeric)
                            {
                                dr[j] = row.GetCell(j).NumericCellValue;
                            }
                            else
                            {
                                dr[j] = row.GetCell(j).ToString().Replace("$", "");
                            }
                        }
                        dt.Rows.Add(dr);
                    }
                }
                fs.Close();
            }
            finally
            {
                if (fs != null) fs.Dispose();
            }
            return dt;
        }

        private void BtnInport_Click(object sender, EventArgs e)
        {
            string thisNewFileName = "";
            GetFileName(out thisNewFileName);
            DataTable thisNewDataTable = GetDataTableFromExcelFile(thisNewFileName);
            DgwResult.DataSource = thisNewDataTable;
        }
    }
}
