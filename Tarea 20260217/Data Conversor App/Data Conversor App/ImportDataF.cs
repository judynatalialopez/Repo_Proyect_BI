using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.Data;
using System.IO;

namespace Data_Conversor_App
{
    public partial class ImportDataF : Form
    {
        public ImportDataF()
        {
            InitializeComponent();
        }

        private void btnCargarExcel_Click(object sender, EventArgs e)
        {
            ofdExcel.Filter = "Excel Files (*.xlsx)|*.xlsx";
            ofdExcel.Title = "Seleccionar archivo Excel";

            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                string ruta = ofdExcel.FileName;
                CargarExcelAGrid(ruta);
            }
        }

        private void CargarExcelAGrid(string rutaArchivo)
        {
            using (var package = new ExcelPackage(new FileInfo(rutaArchivo)))
            {
                ExcelWorksheet hoja = package.Workbook.Worksheets[0];
                DataTable tabla = new DataTable();

                int columnas = hoja.Dimension.End.Column;
                int filas = hoja.Dimension.End.Row;

                for (int col = 1; col <= columnas; col++)
                {
                    tabla.Columns.Add(hoja.Cells[1, col].Text);
                }

                for (int fila = 2; fila <= filas; fila++)
                {
                    DataRow row = tabla.NewRow();

                    for (int col = 1; col <= columnas; col++)
                    {
                        row[col - 1] = hoja.Cells[fila, col].Text;
                    }

                    tabla.Rows.Add(row);
                }

                if (hoja.Dimension == null)
                {
                    MessageBox.Show("El archivo Excel está vacío");
                    return;
                }

                if (string.IsNullOrWhiteSpace(hoja.Cells[1, 1].Text))
                {
                    MessageBox.Show("El Excel no tiene encabezados");
                    return;
                }

                dgvDatos.DataSource = tabla;
            }
        }
    }
}
