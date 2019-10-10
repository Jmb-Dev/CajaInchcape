using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndigo.AppPersistencia.Class.ReportesCaja.Estructura
{
    class RESUMEN_CAJA
    {
        public string ID_SUCURSAL { get; set; }
        public string SUCURSAL { get; set; }
        public string ID_CAJA { get; set; }
        public string NOM_CAJA { get; set; }
        public string MONTO_EFEC { get; set; }
        public string MONTO_DIA { get; set; }
        public string MONTO_FECHA { get; set; }
        public string MONTO_TRANSF { get; set; }
        public string MONTO_VALE_V { get; set; }
        public string MONTO_DEP { get; set; }
        public string MONTO_TARJ { get; set; }
        public string MONTO_FINANC { get; set; }
        public string MONTO_APP { get; set; }
        public string MONTO_CREDITO { get; set; }
        public string MONEDA { get; set; }
        //public string MONTO_C_CURSE { get; set; }
    }
}
