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

                ValidarDatos(tabla); 
                dgvDatos.DataSource = tabla;

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

        private void ImportDataF_Load(object sender, EventArgs e)
        {
            cmbOperador.Items.AddRange(new string[] { "=", ">", "<", "<>", ">=", "<=" });
            cmbOperador.SelectedIndex = 1;
        }

        private void btnAplicarFiltro_Click(object sender, EventArgs e)
        {
            if (dgvDatos.DataSource is DataTable dt)
            {
                string operador = cmbOperador.SelectedItem.ToString();
                string valor = txtValorFiltro.Text;

                if (decimal.TryParse(valor, out decimal salario))
                {
                    DataView dv = dt.DefaultView;
                    dv.RowFilter = $"Salario {operador} {salario}";
                    dgvDatos.DataSource = dv;
                }
                else
                {
                    MessageBox.Show("Ingrese un valor numérico válido.");
                }
            }
        }

        private void btnLimpiarFiltro_Click(object sender, EventArgs e)
        {
            if (dgvDatos.DataSource is DataView dv)
            {
                dv.RowFilter = "";
            }
        }

        private void ValidarDatos(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                // TIPO DOCUMENTO
                string tipo = row["Tipo Documento"]?.ToString().Trim();

                if (string.IsNullOrWhiteSpace(tipo))
                {
                    row.SetColumnError("Tipo Documento", "Obligatorio");
                }
                else if (!(tipo == "CC" || tipo == "TI" || tipo == "CE" || tipo == "PT"))
                {
                    row.SetColumnError("Tipo Documento", "Tipo no válido (CC, TI, CE, PT)");
                }

                // NUMERO DOCUMENTO
                string numero = row["Numero Documento"]?.ToString().Trim();

                if (string.IsNullOrWhiteSpace(numero))
                {
                    row.SetColumnError("Numero Documento", "Obligatorio");
                }
                else if (!long.TryParse(numero, out _) || numero.Length < 6)
                {
                    row.SetColumnError("Numero Documento", "Número inválido");
                }

                // SALARIO
                string salarioTexto = row["Salario"]?.ToString().Trim();

                if (!decimal.TryParse(salarioTexto, out decimal salario) || salario < 0)
                {
                    row.SetColumnError("Salario", "Salario inválido");
                }
            }
        }
    }
}
