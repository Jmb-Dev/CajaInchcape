﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndigo.AppPersistencia.Class.BusquedaAnulacion.Estructura
{
    class DET_COMP
    {
        public string ID_COMPROBANTE { get; set; }
        public string ID_DETALLE { get; set; }
        public string VIA_PAGO { get; set; }
        public string DESCRIP_VP { get; set; }
        public string NUM_CHEQUE { get; set; }
        public string FECHA_VENC { get; set; }
        public string MONTO { get; set; }
        public string MONEDA { get; set; }
        public string NUM_CUOTAS { get; set; }
        public string EMISOR { get; set; }
      
    }
}
