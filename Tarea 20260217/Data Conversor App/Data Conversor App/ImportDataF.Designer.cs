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
            btnAplicarFiltro = new Button();
            btnLimpiarFiltro = new Button();
            ((System.ComponentModel.ISupportInitialize)dgvDatos).BeginInit();
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
            dgvDatos.Size = new Size(1249, 549);
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
            // 
            // btnAplicarFiltro
            // 
            btnAplicarFiltro.Location = new Point(324, 31);
            btnAplicarFiltro.Name = "btnAplicarFiltro";
            btnAplicarFiltro.Size = new Size(94, 29);
            btnAplicarFiltro.TabIndex = 4;
            btnAplicarFiltro.Text = "Aplicar";
            btnAplicarFiltro.UseVisualStyleBackColor = true;
            // 
            // btnLimpiarFiltro
            // 
            btnLimpiarFiltro.Location = new Point(424, 31);
            btnLimpiarFiltro.Name = "btnLimpiarFiltro";
            btnLimpiarFiltro.Size = new Size(94, 29);
            btnLimpiarFiltro.TabIndex = 5;
            btnLimpiarFiltro.Text = "Limpiar";
            btnLimpiarFiltro.UseVisualStyleBackColor = true;
            // 
            // ImportDataF
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1322, 671);
            Controls.Add(btnLimpiarFiltro);
            Controls.Add(btnAplicarFiltro);
            Controls.Add(txtValorFiltro);
            Controls.Add(cmbOperador);
            Controls.Add(btnCargarExcel);
            Controls.Add(dgvDatos);
            Name = "ImportDataF";
            Text = "ImportDataF";
            ((System.ComponentModel.ISupportInitialize)dgvDatos).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView dgvDatos;
        private OpenFileDialog ofdExcel;
        private Button btnCargarExcel;
        private ComboBox cmbOperador;
        private TextBox txtValorFiltro;
        private Button btnAplicarFiltro;
        private Button btnLimpiarFiltro;
    }
}