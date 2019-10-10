using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura
{
    public class DATOS_DOCUMENTOS
    {
        public string TXT_DOCU { get; set; }
        public string NRO_DOCUMENTO { get; set; }
        public string FECHA_DOC { get; set; }
        public string FECHA_VENC_DOC { get; set; }
        public string MONTO_DOC_MO { get; set; }
        public string MONTO_DOC_ML { get; set; }
        public string PEDIDO { get; set; } 
    }
}
