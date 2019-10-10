using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndu.AppPersistencia.Class.RendicionCaja.Estructura
{
    class DETALLE_VP
    {
        public string SOCIEDAD { get; set; }
        public string SOCIEDAD_TXT { get; set; }
        public string ID_COMPROBANTE { get; set; }
        public string ID_DETALLE { get; set; }
        public string VIA_PAGO { get; set; }
        public string MONTO { get; set; }
        public string MONEDA { get; set; }
        public string BANCO { get; set; }
        public string BANCO_TXT { get; set; }
        public string EMISOR { get; set; }
        public string NUM_CHEQUE { get; set; }
        public string COD_AUTORIZACION { get; set; }
        public string CLIENTE { get; set; }
        public string NRO_DOCUMENTO { get; set; }
        public string NUM_CUOTAS { get; set; }
        public string FECHA_VENC { get; set; }
        public string FECHA_EMISION { get; set; }
        public string NOTA_VENTA { get; set; }
        public string TEXTO_POSICION { get; set; }
        public string NULO { get; set; }
        public string NRO_REFERENCIA { get; set; }
        public string TIPO_DOCUMENTO { get; set; }
    }
}
