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
            InitData();
        }

        public void InitData()
        {
            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            WorkBook workbook = WorkBook.Load("D:\\master.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            //Select cells easily in Excel notation and return the calculated value, date, text or formula
            int cellValue = sheet["A2"].IntValue;
            // Read from Ranges of cells elegantly.
            foreach (var cell in sheet["A2:B10"])
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
            }
        }
    }


}
