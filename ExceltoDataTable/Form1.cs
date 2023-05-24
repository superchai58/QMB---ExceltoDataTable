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
using OfficeOpenXml;

namespace ExceltoDataTable
{
    public partial class FrmExcelToDataTable : Form
    {
        public FrmExcelToDataTable()
        {
            InitializeComponent();
        }

        private void FrmExcelToDataTable_Load(object sender, EventArgs e)
        {
            string fileName = "MSD.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"SaveFile\", fileName);
            DataTable dt = new DataTable();

            dt = ExcelDataToDataTable(path);
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("True");
            }
            else
            {
                MessageBox.Show("False");
            }
        }

        public static DataTable ExcelDataToDataTable(string filePath)
        {
            var dt = new DataTable();
            var fi = new FileInfo(filePath);
            // Check if the file exists
            if (!fi.Exists)
                throw new Exception("File " + filePath + " Does Not Exists");

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var xlPackage = new ExcelPackage(fi);
            // get the first worksheet in the workbook
            var worksheet = xlPackage.Workbook.Worksheets["sheet1"];

            dt = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].ToDataTable(c =>
            {
                c.FirstRowIsColumnNames = true;
            });

            return dt;
        }        
    }
}
