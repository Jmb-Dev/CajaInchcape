using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura
{
     public class LOG_APERTURA
    {
        public string MANDT { get; set; }
        public string ID_REGISTRO { get; set; }
        public string LAND { get; set; }
        public string ID_CAJA { get; set; }
        public string USUARIO { get; set; }
        public DateTime FECHA { get; set; }
        public DateTime HORA { get; set; }
        public string MONTO { get; set; }
        public string MONEDA { get; set; }
        public string TIPO_REGISTRO { get; set; }
        public string ID_APERTURA { get; set; }
        public string TXT_CIERRE { get; set; }
        public string BLOQUEO { get; set; }
        public string USUARIO_BLOQ { get; set; }

    }
}
