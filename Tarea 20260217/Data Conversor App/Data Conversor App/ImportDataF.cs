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
using System.Text.RegularExpressions;


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

                if (hoja.Dimension == null)
                {
                    MessageBox.Show("El archivo Excel está vacío");
                    return;
                }

                DataTable tabla = new DataTable();

                int columnas = hoja.Dimension.End.Column;
                int filas = hoja.Dimension.End.Row;

                // Crear columnas
                for (int col = 1; col <= columnas; col++)
                {
                    tabla.Columns.Add(hoja.Cells[1, col].Text);
                }

                // Llenar filas
                for (int fila = 2; fila <= filas; fila++)
                {
                    DataRow row = tabla.NewRow();

                    for (int col = 1; col <= columnas; col++)
                    {
                        row[col - 1] = hoja.Cells[fila, col].Text;
                    }

                    tabla.Rows.Add(row);
                }

                // 🔥 VALIDAR DESPUÉS DE LLENAR
                bool esValido = ValidarDatos(tabla);

                if (!esValido)
                {
                    MessageBox.Show("El archivo contiene celdas vacías o datos inválidos. No se puede cargar.");
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

        private bool ValidarDatos(DataTable dt)
        {
            foreach (DataRow row in dt.Rows)
            {
                // Validar que ninguna celda esté vacía
                foreach (DataColumn col in dt.Columns)
                {
                    if (string.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        return false; // ❌ Bloquea carga
                    }
                }

                // ========= TIPO DOCUMENTO =========
                string tipo = row["Tipo Documento"].ToString().Trim().ToUpper();

                if (!Regex.IsMatch(tipo, @"^[A-Z]+$"))
                    return false;

                if (!(tipo == "CC" || tipo == "TI" || tipo == "CE" || tipo == "PT"))
                    return false;

                // ========= NUMERO DOCUMENTO =========
                string numero = row["Numero Documento"].ToString().Trim();

                if (!Regex.IsMatch(numero, @"^[0-9]{6,}$"))
                    return false;

                // ========= SALARIO =========
                string salarioTexto = row["Sueldo"].ToString().Trim();

                if (!decimal.TryParse(salarioTexto, out decimal salario))
                    return false;

                if (salario <= 0)
                    return false;
            }

            return true; // ✅ Todo válido
        }


        private void txtValorFiltro_TextChanged(object sender, EventArgs e)
        {
            if (dgvDatos.DataSource is DataTable dt)
            {
                if (string.IsNullOrWhiteSpace(txtValorFiltro.Text))
                {
                    dt.DefaultView.RowFilter = "";
                    return;
                }

                if (!decimal.TryParse(txtValorFiltro.Text, out decimal valor))
                    return;

                if (cmbOperador.SelectedItem == null)
                    return;

                string operador = cmbOperador.SelectedItem.ToString();

                dt.DefaultView.RowFilter = $"Convert(Sueldo, 'System.Decimal') {operador} {valor}";
            }
        }


        private void cmbOperador_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbOperador.Items.AddRange(new string[] { "=", ">", "<", "<>", ">=", "<=" });
            cmbOperador.SelectedIndex = 1;
        }

        private void btnLimpiarFiltro_Click_1(object sender, EventArgs e)
        {
            if (dgvDatos.DataSource is DataTable dt)
            {
                dt.DefaultView.RowFilter = "";
            }

            txtValorFiltro.Clear(); // Limpia el campo
            cmbOperador.SelectedIndex = 1; // Vuelve al operador por defecto
        }

    }
}
