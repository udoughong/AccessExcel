using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
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
                    file = new FileStream(@"C:\Users\udoug\OneDrive\文件\MyCase\興利網CASE\tmp\npoi.xls", FileMode.Create);//產生Excel 2003檔案
                    break;
                case ("2007"):
                    file = new FileStream(@"C:\Users\udoug\OneDrive\文件\MyCase\興利網CASE\tmp\npoi.xlsx", FileMode.Create);//產生Excel 2007檔案
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
            string thisInportFileName = "";
            GetFileName(out thisInportFileName);
            DataTable thisInportDataTable = GetDataTableFromExcelFile(thisInportFileName);
            DgwResult.DataSource = thisInportDataTable;
        }

        //範例四，修改Excel(不新增欄位)
        private void UpdateExcelDataWithoutNewCell(string fileName)
        {
            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
                file.Close();
            }

            ISheet sheet = hssfwb.GetSheetAt(0); //抓第1個Sheet工作表.GetSheetAt(0)
            // sheet.SheetName.ToString(); //得到工作表名稱
            //抓合拼欄位的話，只要抓合拼後的最前一個row就好，其他row無法設定值進去
            //例如row5跟row6合拼，只要抓row5就好，row6是無法設值進去
            IRow row = sheet.GetRow(2);//GetRow(2)抓第3個Row

            ICell cell = row.GetCell(1);//GetCell(1)抓第2個Cell
            cell.SetCellValue("test12121213221"); //設定值

            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Write))
            {
                hssfwb.Write(file);
                file.Close();
            }
        }

        private void BtnUpdateExcelWithoutNewCell_Click(object sender, EventArgs e)
        {
            GetFileName(out string thisUpdateFileName);
            UpdateExcelDataWithoutNewCell(thisUpdateFileName);
            DataTable thisNewDataTable = GetDataTableFromExcelFile(thisUpdateFileName);
            DgwResult.DataSource = thisNewDataTable;
        }

        //範例五，修改Excel(新增欄位)
        private void UpdateExcelDataWithNewCell(string fileName)
        {
            #region 把Excel文件載入workbook變數裡，之後關閉Excel文件。
            HSSFWorkbook wk = null;
            using (FileStream fs = File.Open(fileName, FileMode.Open,
            FileAccess.Read, FileShare.ReadWrite))
            {
                wk = new HSSFWorkbook(fs);
                fs.Close();
            }
            #endregion
            #region 新增寫入欄位
            //抓第1個Sheet工作表.GetSheetAt(0)
            ISheet sheet = wk.GetSheetAt(0);
            //因為要新增下一個欄位，所以擷取第一行，建立下一列。
            ICell cell = sheet.GetRow(0).CreateCell(sheet.GetRow(0).Cells.Count);
            ////在第一行，新建立的那一列，寫入欄位名稱。
            cell.SetCellValue("欄位名稱" + sheet.GetRow(0).Cells.Count.ToString());
            #endregion
            #region 存檔
            using (FileStream fileStream
                = File.Open(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                wk.Write(fileStream);
                fileStream.Close();
            }
            #endregion
        }

        private void BtnUpdateExcelWithNewCell_Click(object sender, EventArgs e)
        {
            GetFileName(out string thisUpdateFileName);
            UpdateExcelDataWithNewCell(thisUpdateFileName);
            DataTable thisNewDataTable = GetDataTableFromExcelFile(thisUpdateFileName);
            DgwResult.DataSource = thisNewDataTable;
        }
    }
}
