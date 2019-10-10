using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndu.AppPersistencia.Class.ArqueoCaja.Estructura
{
   public class DETALLE_ARQUEO
    {
        public string LAND { get; set; }
        public string ID_CAJA { get; set; }
        public string USUARIO { get; set; }
        public string SOCIEDAD { get; set; }
        public string FECHA_REND { get; set; }
        public string HORA_REND { get; set; }
        public string MONEDA { get; set; }
        public string VIA_PAGO { get; set; }
        public string TIPO_MONEDA { get; set; }
        public string CANTIDAD_MON { get; set; }
        public string SUMA_MON_BILL { get; set; }
        public string CANTIDAD_DOC { get; set; }
        public string SUMA_DOCS { get; set; }
     }
}
