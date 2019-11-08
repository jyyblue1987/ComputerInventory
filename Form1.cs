using ExcelDataReader;
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


            int row = 2;

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

                            row++;
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
            //xlsxWorkbook.Metadata.Author = "Company";
            ////Add a blank WorkSheet
            //WorkSheet xlsSheet = xlsxWorkbook.CreateWorkSheet("summary");

            ////Add data and styles to the new worksheet
            //xlsSheet["A1"].Value = "Shipment ID";
            //xlsSheet["B1"].Value = "Asset Type";
            //xlsSheet["C1"].Value = "Pallet";
            //xlsSheet["D1"].Value = "Shop";
            //xlsSheet["E1"].Value = "Scrap";
            //xlsSheet["F1"].Value = "At Hand";


            //WorkBook xlsxWorkbook_client = null;
            //WorkSheet xlsSheet_client = null;



            //String prevShipment = "";
            //String ship_id = "";
            //row = 2;
            //foreach( var p in summary)
            //{
            //    if (prevShipment != p.Shipment)
            //    {
            //        ship_id = p.Shipment;

            //        try
            //        {
            //            if(xlsxWorkbook_client != null)
            //            {
            //                xlsxWorkbook_client.SaveAs(destFolder + "\\" + prevShipment + ".xlsx");
            //                xlsxWorkbook_client.Close();
            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //        }

            //        xlsxWorkbook_client = WorkBook.Create(ExcelFileFormat.XLSX);
            //        xlsxWorkbook_client.Metadata.Author = "Company";
            //        //Add a blank WorkSheet
            //        xlsSheet_client = xlsxWorkbook_client.CreateWorkSheet("summary");

            //        //Add data and styles to the new worksheet
            //        xlsSheet_client["A1"].Value = "Shipment ID";
            //        xlsSheet_client["B1"].Value = "Asset Type";
            //        xlsSheet_client["C1"].Value = "Pallet";
            //        xlsSheet_client["D1"].Value = "Shop";
            //        xlsSheet_client["E1"].Value = "Scrap";
            //        xlsSheet_client["F1"].Value = "At Hand";
            //    }
            //    else
            //        ship_id = "";

            //    xlsSheet["A" + row].Value = ship_id;
            //    xlsSheet["B" + row].Value = p.Asset;
            //    xlsSheet["C" + row].Value = p.Pallet;
            //    xlsSheet["D" + row].Value = p.Shop;
            //    xlsSheet["E" + row].Value = p.Scrap;
            //    xlsSheet["F" + row].Value = p.AtHand;

            //    xlsSheet_client["A" + row].Value = ship_id;
            //    xlsSheet_client["B" + row].Value = p.Asset;
            //    xlsSheet_client["C" + row].Value = p.Pallet;
            //    xlsSheet_client["D" + row].Value = p.Shop;
            //    xlsSheet_client["E" + row].Value = p.Scrap;
            //    xlsSheet_client["F" + row].Value = p.AtHand;

            //    prevShipment = p.Shipment;

            //    Console.WriteLine("{0} - {1} - {2} - {3} - {4}", p.Shipment, p.Asset, p.Pallet, p.Shop, p.Scrap, p.AtHand);

            //    row++;
            //}

            //try
            //{
            //    if (xlsxWorkbook_client != null)
            //    {
            //        xlsxWorkbook_client.SaveAs(destFolder + "\\" + prevShipment + ".xlsx");
            //        xlsxWorkbook_client.Close();
            //    }
            //}
            //catch (Exception ex)
            //{
            //}


            ////Save the excel file
            //try
            //{
            //    xlsxWorkbook.SaveAs(destFolder + "\\total_summary.xlsx");
            //    xlsxWorkbook.Close();
            //    txtResult.Text = "Summary is generated!!!";
            //} catch(Exception e)
            //{
            //    txtResult.Text = "Output Directory Error!";
            //}            
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
