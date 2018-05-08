using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.Data;
using System.Data.SqlClient;

namespace Crossexcel
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime date = DateTime.Now;
            string datewithformat = date.ToString();
            string dateday = date.ToString("dd MMMM yyyy HH mm ");
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;
                //IStyle headerStyle = wo.Styles.Add("HeaderStyle");
                IWorkbook workbook = excelEngine.Excel.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.EnableSheetCalculations();
                DataTable tabla = GetExistenciaDataTable();
                int osos = tabla.Rows.Count;

                worksheet.ImportDataTable(tabla, true, 2,1);
                worksheet.AutoFilters.FilterRange = worksheet.Range["A2:F2"];
                worksheet.Range["A1"].Text = "Llantas y Rines del Guadiana S.A. de C.V. - Existencias LRG Al "+dateday+"- B4 Francisco Villa";

                IStyle headerStyle = workbook.Styles.Add("HeaderStyle");
                headerStyle.BeginUpdate();

                workbook.SetPaletteColor(8, System.Drawing.Color.FromArgb(255, 174, 33));

                headerStyle.Color = System.Drawing.Color.FromArgb(255, 174, 33);

                headerStyle.Font.Bold = true;

                headerStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;

                headerStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;

                headerStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;

                headerStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;

                headerStyle.EndUpdate();

                worksheet.Rows[1].CellStyle = headerStyle;

                IStyle pStyle = workbook.Styles.Add("pStyle");
                pStyle.BeginUpdate();

                workbook.SetPaletteColor(9, System.Drawing.Color.FromArgb(59, 187, 219));

                pStyle.Color = System.Drawing.Color.FromArgb(59, 187, 219);

                pStyle.Font.Bold = true;

                pStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;

                pStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;

                pStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;

                pStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;

                pStyle.EndUpdate();

                worksheet.Rows[0].CellStyle = pStyle;
                worksheet.SetColumnWidth(2, 50);

                // Create Table with data in the given range
                int soviet = osos;
                int rojos = soviet + 3;
                int rus = soviet + 4;
                string rusia = rus.ToString();
                string cossacks = rojos.ToString();
                string gulag = "A2:F" + cossacks + "";
                //IListObject table = worksheet.ListObjects.Create("Table1", worksheet[gulag]);
                //IRange range = worksheet.Range[gulag];
                //table.ShowTotals = true;
                //table.Columns[0].TotalsRowLabel = "Total";

                //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                string chorchill = "F2,F" + cossacks + "";
                string russevel = "F" + rusia + "";
                string nrusia = "=SUM(F2:F" + cossacks + ")";
                worksheet.Range[russevel].Formula = nrusia;
                //<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                //table.Columns[5].TotalsCalculation = ExcelTotalsCalculation.Sum;
                //hacer el subtotal pero conformula ** el otro marca error con total calculation 
                //range.SubTotal(0, ConsolidationFunction.Sum, new int[] {1,rojos});
                string namer = dateday;
                string fileName = "LRG-Existencias al " + namer + "B4 Francisco Villa.xlsx";
                workbook.SaveAs(fileName);
                workbook.Close();
                excelEngine.Dispose();

            }


        }
        static DataTable GetExistenciaDataTable()
        {
            string conect = "SERVER = gggctserver.database.windows.net; DATABASE = rdbms_GGGC_Public_TESTING; USER ID = abril; PASSWORD = gggc.2017";

            SqlConnection sqlconn = new SqlConnection(conect);
            sqlconn.Open();
            // string oso = GetTablaSuc(intSUCURSALID);
            //SqlDataAdapter adapter = new SqlDataAdapter(oso, sqlconn);
            SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM dbo.vtaExitencia44  ", sqlconn);
            DataSet dsPubs = new DataSet("Pubs");
            adapter.Fill(dsPubs, "Existencias_Sucursales");
            DataTable dtbl = new DataTable();

            dtbl = dsPubs.Tables["Existencias_Sucursales"];
            sqlconn.Close();
            //tratando de agrgar una columna de un data set ya definido 

            return dtbl;
        }


    }
    
}
