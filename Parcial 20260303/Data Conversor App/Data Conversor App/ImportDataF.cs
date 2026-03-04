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
using Data_Conversor_App.Classes;


namespace Data_Conversor_App
{
    public partial class ImportDataF : Form
    {
        public ImportDataF()
        {
            InitializeComponent();
        }
        private void titulo1_TextChanged(object sender, EventArgs e)
        {

        }

        List<Venta> Ventas = new List<Venta>();
        List<string> errores = new List<string>();


        private void btnCargarExcel_Click(object sender, EventArgs e)
        {
            ofdExcel.Filter = "Excel Files (*.xlsx)|*.xlsx";
            ofdExcel.Title = "Seleccionar archivo Excel";

            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                Ventas = CargarExcelAGrid(ofdExcel.FileName);
                dgvDatos.DataSource = Ventas;

                if (dgvDatos.Columns["TotalVenta"] != null)
                {
                    dgvDatos.Columns["TotalVenta"].DefaultCellStyle.Format = "N0";
                    dgvDatos.Columns["TotalVenta"].DefaultCellStyle.FormatProvider =
                        new System.Globalization.CultureInfo("es-CO");
                }
            }
        }

        private List<Venta> CargarExcelAGrid(string ruta)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(ruta)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];
                int filas = worksheet.Dimension.Rows;

                for (int fila = 4; fila <= filas; fila++)
                {
                    string ciudad = worksheet.Cells[fila, 4].Text.Trim().ToUpper();

                    /*if (ciudad != "Bogota" && ciudad != "Bogotá" && ciudad != "Medellin" && ciudad != "Medellín")

                    {
                        MessageBox.Show($"Ciudad inválida en fila {fila}.");
                        continue;
                    }*/

                    var hoja1 = package.Workbook.Worksheets[0];
                    var hoja3 = package.Workbook.Worksheets[2];
                    string total1 = hoja1.Cells[fila, 8].Text.Trim();
                    string total3 = hoja3.Cells[fila, 8].Text.Trim();
                    if (total1 != total3)

                    {
                        MessageBox.Show($"Error, los valores del Total Venta no son iguales entre sí. {fila}.");
                        continue;
                    }

                    //carga todos 
                    //Venta emp = new Venta 
                    //{ 
                    //    IdVenta = worksheet.Cells[fila, 1].Text, 
                    //    Fecha = worksheet.Cells[fila, 2].Text, 
                    //    TotalVenta = decimal.Parse(worksheet.Cells[fila, 3].Text) 
                    //}; 
                    //Ventas.Add(emp); 

                    //carga solo los ciudads de documentos válidos 
                    Ventas.Add(new Venta

                    {
                        IdVenta = int.Parse(worksheet.Cells[fila, 1].Text),
                        Fecha = DateOnly.Parse(worksheet.Cells[fila, 2].Text),
                        Cliente = worksheet.Cells[fila, 3].Text,
                        Ciudad = worksheet.Cells[fila, 4].Text,
                        Producto = worksheet.Cells[fila, 5].Text,
                        Cantidad = int.Parse(worksheet.Cells[fila, 6].Text),
                        PrecioUnitario = decimal.Parse(worksheet.Cells[fila, 7].Text),
                        TotalVenta = decimal.Parse(worksheet.Cells[fila, 8].Text)
                    });

                    decimal TotalVenta = 0;
                    bool esNumeroValido = decimal.TryParse(

                        worksheet.Cells[fila, 8].Text,

                        System.Globalization.NumberStyles.Any,

                        new System.Globalization.CultureInfo("es-CO"),

                        out TotalVenta

                    );

                    if (!esNumeroValido)
                    {
                        errores.Add($"Fila {fila}: Total Venta inválido");
                        continue;
                    }
                    if (TotalVenta < 0)
                    {
                        errores.Add($"Fila {fila}: Total Venta negativo");
                        continue;
                    }
                }

                if (errores.Any())
                {
                    MessageBox.Show(string.Join("\n", errores));
                }

                decimal totalNomina = Ventas.Sum(e => e.TotalVenta);
                MessageBox.Show(totalNomina.ToString("C0", new System.Globalization.CultureInfo("es-CO")));

                var agrupado = Ventas

                    .GroupBy(e => e.IdVenta)

                    .Select(g => new

                    {
                        IdVenta = g.Key,
                        Cantidad = g.Count(),
                        Total = g.Sum(x => x.TotalVenta)
                    })
                    .ToList();

                dgvResumen.DataSource = agrupado;

                var estadisticas = new
                {

                    Promedio = Ventas.Average(e => e.TotalVenta),
                    Maximo = Ventas.Max(e => e.TotalVenta),
                    Minimo = Ventas.Min(e => e.TotalVenta),
                    Total = Ventas.Sum(e => e.TotalVenta),
                    Cantidad = Ventas.Count()

                };

                dgvEstadisticas.DataSource = new List<object> { estadisticas };

                var duplicados = Ventas
                    .GroupBy(e => e.Fecha)
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

                if (duplicados.Any())

                {
                    MessageBox.Show(
                        "Documentos duplicados:\n" +
                        string.Join("\n", duplicados)
                    );
                }
                else
                {
                    MessageBox.Show("No hay duplicados");
                }
            }
            return Ventas;
        }

        private void btnAplicarFiltro_Click(object sender, EventArgs e)
        {
            if (dgvDatos.DataSource is DataTable dt)
            {
                string operador = cmbOperador.SelectedItem.ToString();
                string valor = txtValorFiltro.Text;

                if (decimal.TryParse(valor, out decimal venta))
                {
                    DataView dv = dt.DefaultView;
                    dv.RowFilter = $"Ventas Totales {operador} {venta}";
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

                foreach (DataColumn col in dt.Columns)
                {
                    if (string.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        return false;
                    }
                }

                string ciudad = row["ciudad Documento"].ToString().Trim().ToUpper();

                if (!Regex.IsMatch(ciudad, @"^[A-Z]+$"))
                    return false;

                if (!(ciudad == "CC" || ciudad == "TI" || ciudad == "CE" || ciudad == "PT"))
                    return false;

                string numero = row["Numero Documento"].ToString().Trim();

                if (!Regex.IsMatch(numero, @"^[0-9]{6,}$"))
                    return false;

                string ventaTexto = row["TotalVenta"].ToString().Trim();

                if (!decimal.TryParse(ventaTexto, out decimal venta))
                    return false;

                if (venta <= 0)
                    return false;
            }

            return true;
        }

        private void ImportDataF_Load(object sender, EventArgs e)
        {
            cmbOperador.Items.AddRange(new string[] { "=", ">", "<", "<>", ">=", "<=" });
            cmbOperador.SelectedIndex = 1;
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

                dt.DefaultView.RowFilter = $"Convert(TotalVenta, 'System.Decimal') {operador} {valor}";
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

            txtValorFiltro.Clear();
            cmbOperador.SelectedIndex = 1;
        }

        //Exportar CVS
        private void ExportarCSV(List<Venta> Ventas)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV (*.csv)|*.csv";
                sfd.FileName = "Nomina.csv";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var lineas = new List<string>();
                    lineas.Add("ciudad_doc,nro_doc,TotalVenta");

                    foreach (var emp in Ventas)
                    {
                        lineas.Add($"{emp.IdVenta},{emp.Fecha},{emp.TotalVenta}");
                    }

                    File.WriteAllLines(sfd.FileName, lineas);

                    MessageBox.Show("Archivo CSV exportado correctamente.");
                }
            }
        }
        private void btnExportar_Click(object sender, EventArgs e)
        {
            if (Ventas != null && Ventas.Any())
            {
                ExportarCSV(Ventas);
            }
            else
            {
                MessageBox.Show("No hay datos para exportar.");
            }
        }
    }
}
