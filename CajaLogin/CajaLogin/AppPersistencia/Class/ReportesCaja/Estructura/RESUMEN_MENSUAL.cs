using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndu.AppPersistencia.Class.ReportesCaja.Estructura
{
    class RESUMEN_MENSUAL
    {
        public string ID_SUCURSAL { get; set; }
        public string ID_CAJA { get; set; }
        public string SUCURSAL { get; set; }
        public string NOM_CAJA { get; set; }
        public string CAJERO { get; set; }
        public string AREA_VTAS { get; set; }
        public string FLUJO_DOCS { get; set; }
        public string TOTAL_MOV { get; set; }
        public string TOTAL_INGR { get; set; }
        public string MONTO_EFEC { get; set; }
        public string MONTO_DIA { get; set; }
        public string MONTO_FECHA { get; set; }
        public string MONTO_TARJ { get; set; }
        public string MONTO_APP { get; set; }
        public string MONTO_TRANSF { get; set; }
        public string MONTO_VALE_V { get; set; }
        public string MONTO_DEP { get; set; }
        public string MONTO_FINANC { get; set; }
        public string MONTO_CREDITO { get; set; }
        //public string MONTO_CREDITO { get; set; }
        public string TOTAL_CAJERO { get; set; }
    }
}
