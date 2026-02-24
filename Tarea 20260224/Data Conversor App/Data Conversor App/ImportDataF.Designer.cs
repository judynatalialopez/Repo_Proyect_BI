namespace Data_Conversor_App
{
    partial class ImportDataF
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            dgvDatos = new DataGridView();
            ofdExcel = new OpenFileDialog();
            btnCargarExcel = new Button();
            cmbOperador = new ComboBox();
            txtValorFiltro = new TextBox();
            btnLimpiarFiltro = new Button();
            dgvResumen = new DataGridView();
            dgvEstadisticas = new DataGridView();
            titulo1 = new TextBox();
            titulo2 = new TextBox();
            btnExportar = new Button();
            ((System.ComponentModel.ISupportInitialize)dgvDatos).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvResumen).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dgvEstadisticas).BeginInit();
            SuspendLayout();
            // 
            // dgvDatos
            // 
            dgvDatos.AllowUserToAddRows = false;
            dgvDatos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvDatos.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvDatos.Location = new Point(33, 91);
            dgvDatos.Name = "dgvDatos";
            dgvDatos.ReadOnly = true;
            dgvDatos.RowHeadersWidth = 51;
            dgvDatos.Size = new Size(1249, 239);
            dgvDatos.TabIndex = 0;
            // 
            // ofdExcel
            // 
            ofdExcel.FileName = "ofdExcel";
            // 
            // btnCargarExcel
            // 
            btnCargarExcel.Location = new Point(1008, 31);
            btnCargarExcel.Name = "btnCargarExcel";
            btnCargarExcel.Size = new Size(274, 29);
            btnCargarExcel.TabIndex = 1;
            btnCargarExcel.Text = "Importar Datos";
            btnCargarExcel.UseVisualStyleBackColor = true;
            btnCargarExcel.Click += btnCargarExcel_Click;
            // 
            // cmbOperador
            // 
            cmbOperador.FormattingEnabled = true;
            cmbOperador.Location = new Point(33, 31);
            cmbOperador.Name = "cmbOperador";
            cmbOperador.Size = new Size(154, 28);
            cmbOperador.TabIndex = 2;
            // 
            // txtValorFiltro
            // 
            txtValorFiltro.Location = new Point(193, 31);
            txtValorFiltro.Name = "txtValorFiltro";
            txtValorFiltro.Size = new Size(125, 27);
            txtValorFiltro.TabIndex = 3;
            txtValorFiltro.TextChanged += txtValorFiltro_TextChanged;
            // 
            // btnLimpiarFiltro
            // 
            btnLimpiarFiltro.Location = new Point(328, 31);
            btnLimpiarFiltro.Name = "btnLimpiarFiltro";
            btnLimpiarFiltro.Size = new Size(94, 29);
            btnLimpiarFiltro.TabIndex = 5;
            btnLimpiarFiltro.Text = "Limpiar";
            btnLimpiarFiltro.UseVisualStyleBackColor = true;
            btnLimpiarFiltro.Click += btnLimpiarFiltro_Click_1;
            // 
            // dgvResumen
            // 
            dgvResumen.AllowUserToAddRows = false;
            dgvResumen.AllowUserToDeleteRows = false;
            dgvResumen.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvResumen.Location = new Point(33, 406);
            dgvResumen.Name = "dgvResumen";
            dgvResumen.ReadOnly = true;
            dgvResumen.RowHeadersWidth = 51;
            dgvResumen.Size = new Size(322, 243);
            dgvResumen.TabIndex = 6;
            // 
            // dgvEstadisticas
            // 
            dgvEstadisticas.AllowUserToAddRows = false;
            dgvEstadisticas.AllowUserToDeleteRows = false;
            dgvEstadisticas.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgvEstadisticas.Location = new Point(385, 406);
            dgvEstadisticas.Name = "dgvEstadisticas";
            dgvEstadisticas.ReadOnly = true;
            dgvEstadisticas.RowHeadersWidth = 51;
            dgvEstadisticas.Size = new Size(897, 243);
            dgvEstadisticas.TabIndex = 7;
            // 
            // titulo1
            // 
            titulo1.Font = new Font("Showcard Gothic", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 0);
            titulo1.Location = new Point(33, 358);
            titulo1.Name = "titulo1";
            titulo1.ReadOnly = true;
            titulo1.Size = new Size(322, 29);
            titulo1.TabIndex = 8;
            titulo1.Text = "AGRUPACIONES POR DOCUMENTOS";
            titulo1.TextAlign = HorizontalAlignment.Center;
            titulo1.TextChanged += titulo1_TextChanged;
            // 
            // titulo2
            // 
            titulo2.Font = new Font("Showcard Gothic", 10.2F, FontStyle.Bold, GraphicsUnit.Point, 0);
            titulo2.Location = new Point(385, 358);
            titulo2.Name = "titulo2";
            titulo2.ReadOnly = true;
            titulo2.Size = new Size(897, 29);
            titulo2.TabIndex = 9;
            titulo2.Text = "ESTADÍSTICAS";
            titulo2.TextAlign = HorizontalAlignment.Center;
            // 
            // btnExportar
            // 
            btnExportar.Location = new Point(706, 31);
            btnExportar.Name = "btnExportar";
            btnExportar.Size = new Size(280, 29);
            btnExportar.TabIndex = 10;
            btnExportar.Text = "Exportar CSV";
            btnExportar.UseVisualStyleBackColor = true;
            btnExportar.Click += btnExportar_Click;
            // 
            // ImportDataF
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1322, 671);
            Controls.Add(btnExportar);
            Controls.Add(titulo2);
            Controls.Add(titulo1);
            Controls.Add(dgvEstadisticas);
            Controls.Add(dgvResumen);
            Controls.Add(btnLimpiarFiltro);
            Controls.Add(txtValorFiltro);
            Controls.Add(cmbOperador);
            Controls.Add(btnCargarExcel);
            Controls.Add(dgvDatos);
            Name = "ImportDataF";
            Text = "ImportDataF";
            Load += ImportDataF_Load;
            ((System.ComponentModel.ISupportInitialize)dgvDatos).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvResumen).EndInit();
            ((System.ComponentModel.ISupportInitialize)dgvEstadisticas).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dgvDatos;
        private OpenFileDialog ofdExcel;
        private Button btnCargarExcel;
        private ComboBox cmbOperador;
        private TextBox txtValorFiltro;
        private Button btnLimpiarFiltro;
        private DataGridView dgvResumen;
        private DataGridView dgvEstadisticas;
        private TextBox titulo1;
        private TextBox titulo2;
        private Button btnExportar;
    }
}