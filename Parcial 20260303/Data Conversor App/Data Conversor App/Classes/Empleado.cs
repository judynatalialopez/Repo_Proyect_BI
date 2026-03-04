using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_Conversor_App.Classes
{
    internal class Venta
    {
        public int IdVenta { get; set; }
        public DateOnly Fecha { get; set; }
        public string Cliente { get; set; }
        public string Ciudad { get; set; }
        public string Producto { get; set; }
        public int Cantidad { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal TotalVenta { get; set; }
    }
}
