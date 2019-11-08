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
        DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();
            InitData();
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

        public void InitData()
        {
            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            WorkBook workbook = WorkBook.Load("D:\\master.xlsx");
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

            DataView view = new DataView(dt);
            DataTable distinctValues = view.ToTable(true, "Shipment");

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
                                Pallet = grp.Sum(r => r.Field<int>("Qty"))
                            };

            foreach( var p in summary)
            {
                Console.WriteLine("{0} - {1} -{2}", p.Shipment, p.Asset, p.Pallet);
            }
            Console.WriteLine(summary);

        }
    }


}
