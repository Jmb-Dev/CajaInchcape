using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndu.AppPersistencia.Class.BusquedaReimpresiones.Estructura
{
   public class DOCUMENTOS
    {
        public string MANDT { get; set; }
        public string LAND { get; set; }
        public string ID_COMPROBANTE { get; set; }
        public string POSICION { get; set; }
        public string CLIENTE { get; set; }
        public string TIPO_DOCUMENTO { get; set; }
        public string SOCIEDAD { get; set; }
        public string NRO_DOCUMENTO { get; set; }
        public string NRO_REFERENCIA { get; set; }
        public string CAJERO_RESP { get; set; }
        public string CAJERO_GEN { get; set; }
        public string ID_CAJA { get; set; }
        public string FECHA_COMP { get; set; }
        public string HORA { get; set; }
        public string NRO_COMPENSACION { get; set; }
        public string TEXTO_CABECERA { get; set; }
        public string NULO { get; set; }
        public string USR_ANULADOR { get; set; }
        public string NRO_ANULACION { get; set; }
        public string APROBADOR_ANULA { get; set; }
        public string TXT_ANULACION { get; set; }
        public string EXCEPCION { get; set; }
        public string FECHA_DOC { get; set; }
        public string FECHA_VENC_DOC { get; set; }
        public string NUM_CUOTA { get; set; }
        public string MONTO_DOC { get; set; }
        public string MONTO_DIFERENCIA { get; set; }
        public string TEXTO_EXCEPCION { get; set; }
        public string PARCIAL { get; set; }
        public string TIME { get; set; }
        public string APROBADOR_EX { get; set; }
        public string MONEDA { get; set; }
        public string CLASE_CUENTA { get; set; }
        public string CLASE_DOC { get; set; }
        public string NUM_CANCELACION { get; set; }
        public string CME { get; set; }
        public string NOTA_VENTA { get; set; }
        public string CEBE { get; set; }
        public string ACC { get; set; }
    }
}
