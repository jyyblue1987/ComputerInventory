using ExcelDataReader;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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


namespace ComputerInventory
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();            
        }

        public static string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }
        public static int NumberFromExcelColumn(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        private String GetString(IExcelDataReader reader, int col)
        {
            var value = reader.GetValue(col);
            if (value == null)
                return "";

            return value.ToString();
        }

        private int GetInt(IExcelDataReader reader, int col)
        {
            var value = reader.GetValue(col);
            if (value == null)
                return 0;

            try
            {
                int result = Int32.Parse(value.ToString());
                return result;
            }
            catch (FormatException)
            {
                return 0;
            }
        }
        public void GenerateSummary(String sourcePath, String destFolder)
        {
            txtResult.Text = "";


            DataTable dt = new DataTable();

            dt.Clear();

            dt.Columns.Add("Asset");
            dt.Columns.Add("Qty", typeof(int));
            dt.Columns.Add("Progress");
            dt.Columns.Add("Shipment");
            
            using (var stream = File.Open(sourcePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 1. Use the reader methods
                    do
                    {
                        reader.Read();
                        while (reader.Read())
                        {
                            String date = GetString(reader, 0);
                            if (date == "")
                                break;

                            String asset_type = GetString(reader, 1);
                            int qty = GetInt(reader, 13);
                            String  progress = GetString(reader, 16);                            
                            String shipment = GetString(reader, 19);                            
                            
                            DataRow record = dt.NewRow();
                            record["Asset"] = asset_type;
                            record["Qty"] = qty;
                            record["Progress"] = progress;
                            record["Shipment"] = shipment.Trim();

                            dt.Rows.Add(record);
                        }
                        break;
                    } while (reader.NextResult());
                }
            }
                      
            
            //DataView view = new DataView(dt);
            //DataTable distinctValues = view.ToTable(true, "Shipment");

            //foreach (DataRow record in distinctValues.Rows)
            //{
            //    Console.WriteLine(record["Shipment"]);               
            //}

            var summary = from rec in dt.AsEnumerable()
                            group rec by new { Shipment = rec.Field<string>("Shipment"), Asset = rec.Field<string>("Asset"), } into grp
                            select new
                            {
                                Shipment = grp.Key.Shipment,
                                Asset = grp.Key.Asset,
                                Pallet = grp.Sum(r => r.Field<int>("Qty") * (r.Field<string>("Progress") == "1" ? 1 : 0) ),
                                Shop = grp.Sum(r => r.Field<int>("Qty") * (r.Field<string>("Progress") == "2" ? 1 : 0)),
                                Scrap = grp.Sum(r => r.Field<int>("Qty") * (r.Field<string>("Progress") == "3" ? 1 : 0)),
                                AtHand = grp.Sum(r => r.Field<int>("Qty") * (r.Field<string>("Progress") == "" ? 1 : 0)),
                            };


            //WorkBook xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("summary");
            IRow row = sheet.CreateRow(0);
            ICell cell = null;
            cell = row.CreateCell(0);
            cell.SetCellValue("Shipment ID");
            cell = row.CreateCell(1);
            cell.SetCellValue("Asset Type");
            cell = row.CreateCell(2);
            cell.SetCellValue("Pallet");
            cell = row.CreateCell(3);
            cell.SetCellValue("Shop");
            cell = row.CreateCell(4);
            cell.SetCellValue("Scrap");
            cell = row.CreateCell(5);
            cell.SetCellValue("At Hand");


            IWorkbook xlsxWorkbook_client = null;
            ISheet xlsSheet_client = null;

            String prevShipment = "";
            String ship_id = "";
            int total_row = 0, client_row = 0;
            foreach (var p in summary)
            {
                if (prevShipment != p.Shipment)
                {
                    ship_id = p.Shipment;

                    try
                    {
                        if (prevShipment != "") 
                        {
                            using (var fs1 = new FileStream(destFolder + "\\" + prevShipment + ".xlsx", FileMode.Create, FileAccess.Write))
                            {
                                xlsxWorkbook_client.Write(fs1);
                                fs1.Close();
                            }
                        }
                    }
                    catch (Exception e)
                    {
                    }

                    xlsxWorkbook_client = new XSSFWorkbook(); ;
                    //Add a blank WorkSheet
                    xlsSheet_client = xlsxWorkbook_client.CreateSheet("summary");

                    //Add data and styles to the new worksheet
                    row = xlsSheet_client.CreateRow(0);
                    cell = row.CreateCell(0);
                    cell.SetCellValue("Shipment ID");
                    cell = row.CreateCell(1);
                    cell.SetCellValue("Asset Type");
                    cell = row.CreateCell(2);
                    cell.SetCellValue("Pallet");
                    cell = row.CreateCell(3);
                    cell.SetCellValue("Shop");
                    cell = row.CreateCell(4);
                    cell.SetCellValue("Scrap");
                    cell = row.CreateCell(5);
                    cell.SetCellValue("At Hand");

                    client_row = 0;

                }
                else
                    ship_id = "";

                client_row++;

                row = xlsSheet_client.CreateRow(client_row);
                cell = row.CreateCell(0);
                cell.SetCellValue(ship_id);
                cell = row.CreateCell(1);
                cell.SetCellValue(p.Asset);
                cell = row.CreateCell(2);
                cell.SetCellValue(p.Pallet);
                cell = row.CreateCell(3);
                cell.SetCellValue(p.Shop);
                cell = row.CreateCell(4);
                cell.SetCellValue(p.Scrap);
                cell = row.CreateCell(5);
                cell.SetCellValue(p.AtHand);

                total_row++;
                row = sheet.CreateRow(total_row);
                cell = row.CreateCell(0);
                cell.SetCellValue(ship_id);
                cell = row.CreateCell(1);
                cell.SetCellValue(p.Asset);
                cell = row.CreateCell(2);
                cell.SetCellValue(p.Pallet);
                cell = row.CreateCell(3);
                cell.SetCellValue(p.Shop);
                cell = row.CreateCell(4);
                cell.SetCellValue(p.Scrap);
                cell = row.CreateCell(5);
                cell.SetCellValue(p.AtHand);

                prevShipment = p.Shipment;

                Console.WriteLine("{0} - {1} - {2} - {3} - {4}", p.Shipment, p.Asset, p.Pallet, p.Shop, p.Scrap, p.AtHand);
            }

            try
            {
                using (var fs1 = new FileStream(destFolder + "\\" + prevShipment + ".xlsx", FileMode.Create, FileAccess.Write))
                {
                    xlsxWorkbook_client.Write(fs1);
                    fs1.Close();
                }
            }
            catch (Exception e)
            {
            }

            using (var fs = new FileStream(destFolder + "\\total_summary.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
                txtResult.Text = "Summary is generated!!!";
            }
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Title = "Select Master File";
            openFileDialog.DefaultExt = "xlsx";
            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = openFileDialog.FileName;
            }
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            folderDialog.ShowNewFolderButton = true;

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                txtOutputPath.Text = folderDialog.SelectedPath;
            }
            
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {            
            GenerateSummary(txtPath.Text, txtOutputPath.Text);
        }
    }


}
