using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IronXL;

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

        public void GenerateSummary(String sourcePath, String destFolder)
        {
            txtResult.Text = ""; 

            WorkBook workbook = null;

            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            try
            {
                workbook = WorkBook.Load(sourcePath);
            } catch(Exception e)
            {
                txtResult.Text = "Source File Error!";
                return;
            }
                        
            WorkSheet sheet = workbook.WorkSheets.First();

            String asset_type_col = "B";
            String qty_col = "N";
            String progress_col = "Q";
            String shipment_col = "T";

            // Find Header
            //foreach (var cell in sheet["A1:Z1"])
            //{
            //    if( cell.Text.ToLower().Contains("asset") )
            //    {
            //        asset_type_col = ExcelColumnFromNumber(cell.ColumnIndex);
            //    }
            //    if (cell.Text.ToLower().Contains("qty"))
            //    {
            //        qty_col = ExcelColumnFromNumber(cell.ColumnIndex);
            //    }
            //    if (cell.Text.ToLower().Contains("progress"))
            //    {
            //        progress_col = ExcelColumnFromNumber(cell.ColumnIndex);
            //    }
            //    if (cell.Text.ToLower().Contains("shipment"))
            //    {
            //        shipment_col = ExcelColumnFromNumber(cell.ColumnIndex);
            //    }
            //}

            //Select cells easily in Excel notation and return the calculated value, date, text or formula
            int cellValue = sheet["A2"].IntValue;

            int row = 2;

            DataTable dt = new DataTable();

            dt.Clear();
            
            dt.Columns.Add("Asset");
            dt.Columns.Add("Qty", typeof(int));
            dt.Columns.Add("Progress");
            dt.Columns.Add("Shipment");

            while (true)
            {
                String date = sheet["A" + row].StringValue;
                if (date == "")
                    break;

                String asset_type = sheet[asset_type_col + row].StringValue;
                int qty = sheet[qty_col + row].IntValue;
                String progress = sheet[progress_col + row].StringValue;
                String shipment = sheet[shipment_col + row].StringValue;
                
                DataRow record = dt.NewRow();
                record["Asset"] = asset_type;
                record["Qty"] = qty;
                record["Progress"] = progress;
                record["Shipment"] = shipment;

                dt.Rows.Add(record);

                row++;
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


            WorkBook xlsxWorkbook = WorkBook.Create(ExcelFileFormat.XLSX);
            xlsxWorkbook.Metadata.Author = "IronXL";
            //Add a blank WorkSheet
            WorkSheet xlsSheet = xlsxWorkbook.CreateWorkSheet("summary");
            //Add data and styles to the new worksheet
            xlsSheet["A1"].Value = "Shipment ID";
            xlsSheet["B1"].Value = "Asset Type";
            xlsSheet["C1"].Value = "Pallet";
            xlsSheet["D1"].Value = "Shop";
            xlsSheet["E1"].Value = "Scrap";
            xlsSheet["F1"].Value = "At Hand";


            String prevShipment = "";
            String ship_id = "";
            row = 2;
            foreach( var p in summary)
            {
                if (prevShipment != p.Shipment)
                    ship_id = p.Shipment;
                else
                    ship_id = "";

                xlsSheet["A" + row].Value = ship_id;
                xlsSheet["B" + row].Value = p.Asset;
                xlsSheet["C" + row].Value = p.Pallet;
                xlsSheet["D" + row].Value = p.Shop;
                xlsSheet["E" + row].Value = p.Scrap;
                xlsSheet["F" + row].Value = p.AtHand;

                prevShipment = p.Shipment;

                Console.WriteLine("{0} - {1} - {2} - {3} - {4}", p.Shipment, p.Asset, p.Pallet, p.Shop, p.Scrap, p.AtHand);

                row++;
            }

            //Save the excel file
            try
            {
                xlsxWorkbook.SaveAs(destFolder + "\\total_summary.xlsx");
                txtResult.Text = "Summary is generated!!!";
            } catch(Exception e)
            {
                txtResult.Text = "Output Directory Error!";
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
