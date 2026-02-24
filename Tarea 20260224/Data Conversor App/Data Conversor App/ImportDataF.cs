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

        List<Empleado> empleados = new List<Empleado>();
        List<string> errores = new List<string>();


        private void btnCargarExcel_Click(object sender, EventArgs e)
        {
            ofdExcel.Filter = "Excel Files (*.xlsx)|*.xlsx";
            ofdExcel.Title = "Seleccionar archivo Excel";

            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                empleados = CargarExcelAGrid(ofdExcel.FileName);
                dgvDatos.DataSource = empleados;

                if (dgvDatos.Columns["Sueldo"] != null)
                {
                    dgvDatos.Columns["Sueldo"].DefaultCellStyle.Format = "N0";
                    dgvDatos.Columns["Sueldo"].DefaultCellStyle.FormatProvider =
                        new System.Globalization.CultureInfo("es-CO");
                }
            }
        }

        private List<Empleado> CargarExcelAGrid(string ruta)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(ruta)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int filas = worksheet.Dimension.Rows;

                for (int fila = 2; fila <= filas; fila++)
                {
                    string tipo = worksheet.Cells[fila, 1].Text.Trim().ToUpper();

                    if (tipo != "CC" && tipo != "CE")

                    {
                        MessageBox.Show($"Tipo de documento inválido en fila {fila}. Solo se permite CC o CE.");
                        continue;
                    }

                    //carga todos 
                    //Empleado emp = new Empleado 
                    //{ 
                    //    TipoDoc = worksheet.Cells[fila, 1].Text, 
                    //    NroDoc = worksheet.Cells[fila, 2].Text, 
                    //    Sueldo = decimal.Parse(worksheet.Cells[fila, 3].Text) 
                    //}; 
                    //empleados.Add(emp); 

                    //carga solo los tipos de documentos válidos 
                    empleados.Add(new Empleado

                    {
                        TipoDoc = tipo,
                        NroDoc = worksheet.Cells[fila, 2].Text,
                        Sueldo = decimal.Parse(worksheet.Cells[fila, 3].Text)
                    });

                    decimal sueldo = 0;
                    bool esNumeroValido = decimal.TryParse(

                        worksheet.Cells[fila, 3].Text,

                        System.Globalization.NumberStyles.Any,

                        new System.Globalization.CultureInfo("es-CO"),

                        out sueldo

                    );

                    if (!esNumeroValido)
                    {
                        errores.Add($"Fila {fila}: Sueldo inválido");
                        continue;
                    }
                    if (sueldo < 0)
                    {
                        errores.Add($"Fila {fila}: Sueldo negativo");
                        continue;
                    }
                }

                if (errores.Any())
                {
                    MessageBox.Show(string.Join("\n", errores));
                }

                decimal totalNomina = empleados.Sum(e => e.Sueldo);
                MessageBox.Show(totalNomina.ToString("C0", new System.Globalization.CultureInfo("es-CO")));

                var agrupado = empleados

                    .GroupBy(e => e.TipoDoc)

                    .Select(g => new

                    {
                        TipoDoc = g.Key,
                        Cantidad = g.Count(),
                        Total = g.Sum(x => x.Sueldo)
                    })
                    .ToList();

                dgvResumen.DataSource = agrupado;

                var estadisticas = new
                {

                    Promedio = empleados.Average(e => e.Sueldo),
                    Maximo = empleados.Max(e => e.Sueldo),
                    Minimo = empleados.Min(e => e.Sueldo),
                    Total = empleados.Sum(e => e.Sueldo),
                    Cantidad = empleados.Count()

                };

                dgvEstadisticas.DataSource = new List<object> { estadisticas };

                var duplicados = empleados
                    .GroupBy(e => e.NroDoc)
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
            return empleados;
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

                foreach (DataColumn col in dt.Columns)
                {
                    if (string.IsNullOrWhiteSpace(row[col].ToString()))
                    {
                        return false;
                    }
                }

                string tipo = row["Tipo Documento"].ToString().Trim().ToUpper();

                if (!Regex.IsMatch(tipo, @"^[A-Z]+$"))
                    return false;

                if (!(tipo == "CC" || tipo == "TI" || tipo == "CE" || tipo == "PT"))
                    return false;

                string numero = row["Numero Documento"].ToString().Trim();

                if (!Regex.IsMatch(numero, @"^[0-9]{6,}$"))
                    return false;

                string salarioTexto = row["Sueldo"].ToString().Trim();

                if (!decimal.TryParse(salarioTexto, out decimal salario))
                    return false;

                if (salario <= 0)
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

            txtValorFiltro.Clear();
            cmbOperador.SelectedIndex = 1;
        }

        //Exportar CVS
        private void ExportarCSV(List<Empleado> empleados)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV (*.csv)|*.csv";
                sfd.FileName = "Nomina.csv";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var lineas = new List<string>();
                    lineas.Add("tipo_doc,nro_doc,sueldo");

                    foreach (var emp in empleados)
                    {
                        lineas.Add($"{emp.TipoDoc},{emp.NroDoc},{emp.Sueldo}");
                    }

                    File.WriteAllLines(sfd.FileName, lineas);

                    MessageBox.Show("Archivo CSV exportado correctamente.");
                }
            }
        }
        private void btnExportar_Click(object sender, EventArgs e)
        {
            if (empleados != null && empleados.Any())
            {
                ExportarCSV(empleados);
            }
            else
            {
                MessageBox.Show("No hay datos para exportar.");
            }
        }
    }
}
