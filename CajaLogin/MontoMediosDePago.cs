using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CajaIndigo
{
    public class VIAS_PAGOGDAUX
    {
        public bool ISSELECTED { get; set; }
        public string SELECCION { get; set; }
        public string ID_CAJA { get; set; }
        public string ID_APERTURA { get; set; }
        public string ID_CIERRE { get; set; }
        public string TEXT_VIA_PAGO { get; set; }
        public string FECHA_EMISION { get; set; }
        public string NUM_DOC { get; set; }
        public string TEXT_BANCO { get; set; }
        public string MONTO_DOC { get; set; }
        public string ZUONR { get; set; }
        public string FECHA_VENC { get; set; }
        public string MONEDA { get; set; }
        public string ID_BANCO { get; set; }
        public string VIA_PAGO { get; set; }
        public string NUM_DEPOSITO { get; set; }
        public string USUARIO { get; set; }
        public string ID_DEPOSITO { get; set; }
        public string FEC_DEPOSITO { get; set; }
        public string BANCO { get; set; }
        public string CTA_BANCO { get; set; }
        public string BELNR_DEP { get; set; }
        public string BELNR { get; set; }
        public string SOCIEDAD { get; set; }
        public string HKONT { get; set; }
        public string ID_COMPROBANTE { get; set; }
        public string ID_DETALLE { get; set; }
    }
    public class DOCUMENTOSAUX
    {
        public bool ISSELECTED { get; set; }
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

    public class VIAS_PAGO_MASIVO
    {
        public string MANDT { get; set; }
        public string LAND { get; set; }
        public string ID_COMPROBANTE { get; set;}
        public string ID_DETALLE { get; set; }
        public string ID_CAJA { get; set; }
        public string VIA_PAGO { get; set; }
        public string MONTO { get; set; }
        public string MONEDA { get; set;}
        public string BANCO { get; set; }
        public string EMISOR { get; set;}
        public string NUM_CHEQUE { get; set; }
        public string COD_AUTORIZACION { get; set;}
        public string NUM_CUOTAS { get; set;}
        public string FECHA_VENC { get; set; }
        public string TEXTO_POSICION { get; set; }
        public string ANEXO { get; set; }
        public string SUCURSAL { get; set; }
        public string NUM_CUENTA { get; set;}
        public string NUM_TARJETA { get; set; }
        public string NUM_VENTA { get; set; }
        public string PAGARE { get; set; }
        public string NUM_VALE_VISTA { get; set; }
        public string PATENTE { get; set; }
        public string FECHA_EMISION { get; set; }
        public string NOMBRE_GIRADOR { get; set; }
        public string CARTA_CURSE { get; set;}
        public string NUM_TRANSFER { get; set;}
        public string NUM_DEPOSITO { get; set;}
        public string CTA_BANCO { get; set;}
        public string IFINAN { get; set;}
        public string ZUONR { get; set; }
        public string CORRE { get; set; }
        public string HKONT { get; set; }
        public string PRCTR { get; set; }
        public string ZNOP  { get; set; }

        public VIAS_PAGO_MASIVO(string mandt, string land, string id_comprobante, string id_detalle, string id_caja, string via_pago, double monto
            , string moneda, string banco, string emisor, string num_cheque, string cod_autorizacion, string num_cuotas, string fecha_venc
            , string texto_posicion, string anexo, string sucursal, string num_cuenta, string num_tarjeta, string num_vale_vista
            , string patente,string num_venta, string pagare,  string fecha_emision, string nombre_girador, string carta_curse
            , string num_transfer, string num_deposito, string cta_banco, string ifinan, string corre, string zuonr, string hkont, string prctr, string znop)
        {
            this.MANDT = mandt;
            this.LAND = land;
            this.ID_COMPROBANTE = id_comprobante;
            this.ID_DETALLE = id_detalle;
            this.VIA_PAGO = via_pago;
            this.ID_CAJA = id_caja;
            this.MONTO = Convert.ToString(monto);
            this.MONEDA = moneda;
            this.BANCO = banco;
            this.EMISOR = emisor;
            this.NUM_CHEQUE = num_cheque;
            this.COD_AUTORIZACION = cod_autorizacion;
            this.NUM_CUOTAS = num_cuotas;
            this.FECHA_VENC = fecha_venc;
            this.TEXTO_POSICION = texto_posicion;
            this.ANEXO = anexo;
            this.SUCURSAL = sucursal;
            this.NUM_CUENTA = num_cuenta;
            this.NUM_TARJETA = num_tarjeta;
            this.NUM_VALE_VISTA = num_vale_vista;
            this.PATENTE = patente;
            this.NUM_VENTA = num_venta;
            this.PAGARE = pagare;
            this.FECHA_EMISION = fecha_emision;
            this.NOMBRE_GIRADOR = nombre_girador;
            this.CARTA_CURSE = carta_curse;
            this.NUM_TRANSFER = num_transfer;
            this.NUM_DEPOSITO = num_deposito;
            this.CTA_BANCO = cta_banco;
            this.IFINAN = ifinan;
            this.CORRE = corre;
            this.ZUONR = zuonr;
            this.HKONT = hkont;
            this.PRCTR = prctr;
            this.ZNOP = znop;
            //this.CODAUT = codaut;
            //this.CODOP = codop;
            //this.ASIG = asig;

           }
       
    }


    public class T_DOCUMENTOS_AUX
    {
        public bool ISSELECTED { get; set; }
        public string NDOCTO { get; set; }
        public string MONTOF { get; set; }
        public string MONTO { get; set; }
        public string MONEDA { get; set; }
        public string FECVENCI { get; set; }
        public string CONTROL_CREDITO { get; set; }
        public string CEBE { get; set; }
        public string COND_PAGO { get; set; }
        public string RUTCLI { get; set; }
        public string NOMCLI { get; set; }
        public string ESTADO { get; set; }
        public string ICONO { get; set; }
        public string DIAS_ATRASO { get; set; }
        public string MONTO_ABONADO { get; set; }
        public string MONTOF_ABON { get; set; }
        public string MONTO_PAGAR { get; set; }
        public string MONTOF_PAGAR { get; set; }
        public string NREF { get; set; }
        public string FECHA_DOC { get; set; }
        public string COD_CLIENTE { get; set; }
        public string SOCIEDAD { get; set; }
        public string CLASE_DOC { get; set; }
        public string CLASE_CUENTA { get; set; }
        public string CME { get; set; }
        public string ACC { get; set; }
        public string FACT_SD_ORIGEN { get; set; }
        public string FACT_ELECT { get; set; }
        public string ID_COMPROBANTE { get; set; }
        public string ID_CAJA { get; set; }
        public string LAND { get; set; }
        public string BAPI { get; set; }
    }
    class CAB_COMPAUX
    {
        public bool ISSELECTED { get; set; }
        public string LAND { get; set; }
        public string ID_CAJA { get; set; }
        public string ID_COMPROBANTE { get; set; }
        public string TIPO_DOCUMENTO { get; set; }
        public string DESCRIPCION { get; set; }
        public string NRO_REFERENCIA { get; set; }
        public string FECHA_COMP { get; set; }
        public string FECHA_VENC_DOC { get; set; }
        public string MONTO_DOC { get; set; }
        public string TEXTO_EXCEPCION { get; set; }
        public string CLIENTE { get; set; }
        public string MONEDA { get; set; }
        public string CLASE_DOC { get; set; }
        public string TXT_CLASE_DOC { get; set; }
        public string NUM_CANCELACION { get; set; }
        public string AUT_JEF { get; set; }

        public string VIA_PAGO { get; set; }

    }
    public class T_DOCUMENTOSAUX
    {
        public string ID { get; set; }
        public bool ISSELECTED { get; set; }
        public string NDOCTO { get; set; }
        public string MONTOF { get; set; }
        public string MONTO { get; set; }
        public string MONEDA { get; set; }
        public string FECVENCI { get; set; }
        public string CONTROL_CREDITO { get; set; }
        public string CEBE { get; set; }
        public string COND_PAGO { get; set; }
        public string RUTCLI { get; set; }
        public string NOMCLI { get; set; }
        public string ESTADO { get; set; }
        public string ICONO { get; set; }
        public string DIAS_ATRASO { get; set; }
        public string MONTO_ABONADO { get; set; }
        public string MONTOF_ABON { get; set; }
        public string MONTO_PAGAR { get; set; }
        public string MONTOF_PAGAR { get; set; }
        public string NREF { get; set; }
        public string FECHA_DOC { get; set; }
        public string COD_CLIENTE { get; set; }
        public string SOCIEDAD { get; set; }
        public string CLASE_DOC { get; set; }
        public string CLASE_CUENTA { get; set; }
        public string CME { get; set; }
        public string ACC { get; set; }
    }

    public class DTE_SII
    {
        public bool ISSELECTED { get; set; }
        public string VBELN { get; set; }
        public string KONDA { get; set; }
        public string BUKRS { get; set; }
        public string XBLNR { get; set; }
        public string ZUONR { get; set; }
        public string TDSII { get; set; }
        public string FODOC { get; set; }
        public string WAERS { get; set; }
        public string FECIMP { get; set; }
        public string HORIM { get; set; }
        public string URLSII { get; set; }
    }

    public class IT_PAGOSAUX
    {
        public bool ISSELECTED { get; set; }
        public string VBELN { get; set; }
        public string CORRE { get; set; }
        public string VIADP { get; set; }
        public string DESCV { get; set; }
        public string DBM_LICEXT { get; set; }
        public string NUDOC { get; set; }
        public string CODBA { get; set; }
        public string NOMBA { get; set; }
        public string CODIN { get; set; }
        public string NOMIN { get; set; }
        public string KUNNR { get; set; }
        public string MONTO { get; set; }
        public string CTACE { get; set; }
        public string FEACT { get; set; }
        public string FEVEN { get; set; }
        public string INTER { get; set; }
        public string TASAI { get; set; }
        public string CUOTA { get; set; }
        public string MINTE { get; set; }
        public string TOTIN { get; set; }
        public string RUTGI { get; set; }
        public string NOMGI { get; set; }
        public string WAERS { get; set; }
        public string STAT { get; set; }
        public string PRCTR { get; set; }
        public string KKBER { get; set; }
        public string STCD1 { get; set; }
        public string BANKN { get; set; }
        public string HKONT { get; set; }

    }
   

    public class AutorizacionViasPago
    {
        public string VBELN { get; set; }
        public string VIADP { get; set; }
        public string DESCV { get; set; }
        public string AUTORIZACION { get; set; }
        public string NUMTARJETA { get; set; }
        public string OPERACION { get; set; }
        public string ASIGNACION { get; set; }
        public string FEC_EMISION { get; set; }

        public AutorizacionViasPago(string vbeln, string mediopago,string descmediopag, string numtarjeta, string autorizacion
            , string operacion, string asignacion, string fec_emision)
        {
            this.VBELN = vbeln;
            this.VIADP = mediopago;
            this.DESCV = descmediopag;
            this.NUMTARJETA = numtarjeta;
            this.AUTORIZACION = autorizacion;
            this.OPERACION = operacion;
            this.ASIGNACION = asignacion;
            this.FEC_EMISION = fec_emision;
            
        }
    }
   public class MontoMediosdePago
    {
        
        public string MedioPago { get; set; }
        public string Monto { get; set; }

        public MontoMediosdePago(string mediopago, string monto)
        {
            this.MedioPago = mediopago;
            this.Monto = monto;
        }
    }

   public class PagosMasivosNuevo
   {

       public string ROW { get; set; }
       public string COL { get; set; }
       public string VALUE { get; set; }

       public PagosMasivosNuevo(string row, string col, string value)
       {
           this.ROW = row;
           this.COL = col;
           this.VALUE = value;
         
       }

       public PagosMasivosNuevo()
       {
           // TODO: Complete member initialization
       }
   }

   public class PagosMasivos
   {

       public string Referencia { get; set; }
       public string Monto { get; set; }
       public string Moneda { get; set; }

       public PagosMasivos(string referencia, string monto, string moneda)
       {
           this.Referencia = referencia;
           this.Monto = monto;
           this.Moneda = moneda;
       }

       public PagosMasivos()
       {
           // TODO: Complete member initialization
       }
   }

   public class ViasPago
   {

       public string Acc { get; set; }
       public string Cond_Pago { get; set; }
       public string Caja { get; set; }

       public ViasPago(string acc, string cond_pago, string caja)
       {
           this.Acc = acc;
           this.Cond_Pago = cond_pago;
           this.Caja = caja;
       }
   }

   class DetalleDocumentosPago 
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
       public string APROBADOR_OK { get; set; }
       public string MONEDA { get; set; }
       public string CLASE_CUENTA { get; set; }
       public string CLASE_DOC { get; set; }
       public string NUM_CANCELACION { get; set; }
       public string CME { get; set; }
       public string NOTA_VENTA { get; set; }
       public string CEBE { get; set; }
       public string ACC { get; set; }

       public DetalleDocumentosPago(string mandt, string land, string id_comprobante, string posicion, string cliente, string tipo_documento
           , string sociedad, string nro_documento, string nro_referencia, string cajero_resp, string cajero_gen, string id_caja, string fecha_comp
           , string hora, string nro_compesacion, string texto_cabecera, string nulo, string usr_anulador, string nro_anulacion
           , string aprobador_anula, string txt_anulacion, string excepcion, string fecha_doc, string fecha_venc_doc, string num_cuota
           , string monto_doc, string monto_diferencia, string texto_excepcion, string parcial, string time, string aprobador_ex, string moneda
           , string clase_cuenta, string clase_doc, string num_cancelacion, string cme, string nota_venta, string cebe, string acc)
       {
           this.MANDT = mandt;
           this.LAND = land;
           this.ID_COMPROBANTE = id_comprobante;
           this.POSICION = posicion;
           this.CLIENTE = cliente;
           this.TIPO_DOCUMENTO = tipo_documento;
           this.SOCIEDAD = sociedad;
           this.NRO_DOCUMENTO = nro_documento;
           this.NRO_REFERENCIA= nro_referencia;
           this.CAJERO_RESP = cajero_resp;
           this.CAJERO_GEN = cajero_gen;
           this.ID_CAJA = id_caja;
           this.FECHA_COMP = fecha_comp;
           this.HORA = hora;
           this.NRO_COMPENSACION = nro_compesacion;
           this.TEXTO_CABECERA = texto_cabecera;
           this.NULO = nulo;
           this.USR_ANULADOR = usr_anulador;
           this.NRO_ANULACION = nro_anulacion;
           this.APROBADOR_ANULA = aprobador_anula;
           this.TXT_ANULACION = txt_anulacion;
           this.EXCEPCION= excepcion;
           this.FECHA_DOC = fecha_doc;
           this.FECHA_VENC_DOC = fecha_venc_doc;
           this.NUM_CUOTA = num_cuota;
           this.MONTO_DOC = monto_doc;
           this.MONTO_DIFERENCIA = monto_diferencia;
           this.TEXTO_EXCEPCION = texto_excepcion;
           this.PARCIAL = parcial;
           this.TIME = time;
           this.APROBADOR_OK = aprobador_ex;
           this.MONEDA = moneda;
           this.CLASE_CUENTA = clase_cuenta;
           this.CLASE_DOC = clase_doc;
           this.NUM_CANCELACION = num_cancelacion;
           this.CME = cme;
           this.NOTA_VENTA = nota_venta;
           this.CEBE = cebe;
           this.ACC = acc;

       }
   }

    class DetalleViasPago
    {
 
        public string MANDT { get; set; }
        public string LAND { get; set; }
        public string ID_COMPROBANTE { get; set; }
        public string ID_DETALLE { get; set; }
        public string VIA_PAGO { get; set; }
        public double MONTO { get; set; }
        public string MONEDA { get; set; }
        public string BANCO { get; set; }
        public string EMISOR { get; set; }
        public string NUM_CHEQUE { get; set; }
        public string COD_AUTORIZACION { get; set; }
        public string NUM_CUOTAS { get; set; }
        public string FECHA_VENC { get; set; }
        public string TEXTO_POSICION { get; set; }
        public string ANEXO { get; set; }
        public string SUCURSAL { get; set; }
        public string NUM_CUENTA { get; set; }
        public string NUM_TARJETA { get; set; }
        public string NUM_VALE_VISTA { get; set; }
        public string PATENTE { get; set; }
        public string NUM_VENTA { get; set; }
        public string PAGARE { get; set; } 
        public string FECHA_EMISION { get; set; }
        public string NOMBRE_GIRADOR { get; set; }
        public string CARTA_CURSE { get; set; }
        public string NUM_TRANSFER { get; set; }
        public string NUM_DEPOSITO { get; set; }
        public string CTA_BANCO { get; set; }
        public string IFINAN { get; set; }
        public string CORRE { get; set; }
        public string ZUONR { get; set; }
        public string HKONT { get; set; }
        public string PRCTR { get; set; }
        public string ZNOP { get; set; } 

        public DetalleViasPago()
        {

        }
        public DetalleViasPago(string mandt, string land, string id_comprobante, string id_detalle, string via_pago, double monto
            , string moneda, string banco, string emisor, string num_cheque, string cod_autorizacion, string num_cuotas, string fecha_venc
            , string texto_posicion, string anexo, string sucursal, string num_cuenta, string num_tarjeta, string num_vale_vista
            , string patente,string num_venta, string pagare,  string fecha_emision, string nombre_girador, string carta_curse
            , string num_transfer, string num_deposito, string cta_banco, string ifinan, string corre, string zuonr, string hkont, string prctr, string znop)
        {
            this.MANDT = mandt;
            this.LAND = land;
            this.ID_COMPROBANTE = id_comprobante;
            this.ID_DETALLE = id_detalle;
            this.VIA_PAGO = via_pago;
            this.MONTO = monto;
            this.MONEDA = moneda;
            this.BANCO = banco;
            this.EMISOR = emisor;
            this.NUM_CHEQUE = num_cheque;
            this.COD_AUTORIZACION = cod_autorizacion;
            this.NUM_CUOTAS = num_cuotas;
            this.FECHA_VENC = fecha_venc;
            this.TEXTO_POSICION = texto_posicion;
            this.ANEXO = anexo;
            this.SUCURSAL = sucursal;
            this.NUM_CUENTA = num_cuenta;
            this.NUM_TARJETA = num_tarjeta;
            this.NUM_VALE_VISTA = num_vale_vista;
            this.PATENTE = patente;
            this.NUM_VENTA = num_venta;
            this.PAGARE = pagare;
            this.FECHA_EMISION = fecha_emision;
            this.NOMBRE_GIRADOR = nombre_girador;
            this.CARTA_CURSE = carta_curse;
            this.NUM_TRANSFER = num_transfer;
            this.NUM_DEPOSITO = num_deposito;
            this.CTA_BANCO = cta_banco;
            this.IFINAN = ifinan;
            this.CORRE = corre;
            this.ZUONR = zuonr;
            this.HKONT = hkont;
            this.PRCTR = prctr;
            this.ZNOP = znop;
            //this.CODAUT = codaut;
            //this.CODOP = codop;
            //this.ASIG = asig;          
        }
    }

    class DetalleDocs
    {
        public string NumDoc { get; set; }
        public string Referenc { get; set; }
        public string RUT { get; set; }
        public string Codigo { get; set; }
        public string Nombre { get; set; }
        public string CeBe { get; set; }
        public string FechaDoc { get; set; }
        public string FechaVenc { get; set; }
        public int DiasRetr { get; set; }
        public String Estado { get; set; }
        public String Moneda { get; set; }
        public Int64 Monto { get; set; }
        public Int64 MontoPend { get; set; }
        public Int64 Saldo { get; set; }

        public DetalleDocs(string numdoc,string referenc, string rut, string codigo, string nombre, string cebe,
            string fechadoc, string fechavenc, int diasretr, string estado, string moneda, 
            Int64 monto, Int64 montopend, Int64 saldo)
        {
            this.NumDoc = numdoc;
            this.Referenc = referenc;
            this.RUT = rut;
            this.Codigo = Codigo;
            this.Nombre = nombre;
            this.CeBe = cebe;
            this.FechaDoc = fechadoc;
            this.FechaVenc = FechaVenc;
            this.DiasRetr = diasretr;
            this.Estado = estado;
            this.Moneda = moneda;
            this.Monto = monto;
            this.MontoPend = montopend;
            this.Saldo = saldo;
        }
    }

    class DetalleMonitor
    {
        public string Sociedad { get; set; }
        public string RUT { get; set; }
       // public string Monto { get; set; }
       // public string NumCheque { get; set; }
        public Int64 Monto { get; set; }

        public DetalleMonitor(string sociedad, string rut, Int64 monto)
        {
            this.Sociedad = sociedad;
            this.RUT = rut;
            //this.Nombre = nombre;
            //this.NumCheque = numchq;
            this.Monto = monto;
        }
    }

    class EsSeleccionada
    {
        public bool IsSelected { get; set; }
        
        public EsSeleccionada(bool isselected)
        {
            this.IsSelected = isselected;
        }
    }

     public class FormatoMonedas
        {
         public string FormatoMonedaCaja(string Valor, string Origen, string Parametro)
         {
             string ValorFormateado = "";
             try
             {
                 
                 if (Origen == "Ch")
                 {
                     if (Valor.Contains("-"))
                     {
                         Valor = "-" + Valor.Replace("-", "");
                     }
                     Valor = Valor.Replace(".", "");
                     Valor = Valor.Replace(",", "");
                     decimal ValorAux;
                     if (Parametro == "2")
                         if (Valor == "0")
                         {
                             ValorAux = Convert.ToDecimal(Valor);
                         }
                         else
                         {
                            ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                         }
                     else
                         ValorAux = Convert.ToDecimal(Valor);
                     ValorFormateado = string.Format("{0:0,0}", ValorAux);

                 }
                 if (Origen == "Ex")
                 {
                     decimal ValorAux = Convert.ToDecimal(Valor);
                     ValorFormateado = string.Format("{0:0,0.##}", ValorAux);
                 }
                 
             }
             catch (Exception ex)
             {
                 Console.WriteLine(ex.Message + ex.StackTrace);
                 System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

             }
             return ValorFormateado;
             GC.Collect();
         }

            public string FormatoMonedaChilena(string Valor, string a)
            {
                string ValorFormateado = "";
                if (Valor.Contains("-"))
                {
                    Valor = "-" + Valor.Replace("-", "");
                }
                Valor = Valor.Replace(".", "");
                Valor = Valor.Replace(",", "");
                decimal ValorAux = 0;
                if (Valor == "0")
                {
                    ValorAux = Convert.ToDecimal(Valor);
                }
                else
                {
                    if (a == "2")
                    {
                        ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                    }
                    if (a == "1")
                    {
                        ValorAux = Convert.ToDecimal(Valor);
                    }
                }
               
                ValorFormateado = string.Format("{0:0,0}", ValorAux);
                return ValorFormateado;
                GC.Collect();
            }

            public string FormatoMonedaChilena2(string Valor, string a)
            {
                string ValorFormateado = "";
                if (Valor.Contains("-"))
                {
                    Valor = "-" + Valor.Replace("-", "");
                }
                Valor = Valor.Replace(".", "");
                Valor = Valor.Replace(",", "");
                decimal ValorAux = 0;
                if (Valor == "0")
                {
                    ValorAux = Convert.ToDecimal(Valor);
                }
                else
                {
                    if (a == "2")
                    {
                        ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                    }
                    if (a == "1")
                    {
                        ValorAux = Convert.ToDecimal(Valor);
                    }
                }
                string ValorAux2 = Convert.ToString(ValorAux);
                ValorFormateado = string.Format("{0:0,0}", ValorAux2);
                //}
                return ValorFormateado;
                GC.Collect();
            }

            public string FormatoMoneda(string Valor)
            {
                string ValorFormateado = string.Empty;
                ValorFormateado = Valor;
                //ValorFormateado = ValorFormateado.Replace(",", ".");
                decimal ValorAux = Convert.ToDecimal(ValorFormateado);
                ValorFormateado = string.Format("{0:0,0}", ValorAux);
                return ValorFormateado;
                GC.Collect();
            }

            public string FormatoMoneda2(string Valor)
            {
                string ValorFormateado = string.Empty;
                ValorFormateado = Valor;
                //ValorFormateado = ValorFormateado.Replace(",", ".");
                decimal ValorAux = Convert.ToDecimal(ValorFormateado);
                ValorFormateado = string.Format("{0:0.##}", ValorAux);
                return ValorFormateado;
                GC.Collect();
            }
            public string FormatoMonedaExtranjera(string Valor)
            {
                string ValorFormateado = "";
                decimal ValorAux = Convert.ToDecimal(Valor);
                ValorFormateado = string.Format("{0:0,0.##}", ValorAux);
                return ValorFormateado;
                GC.Collect();
            }
        }
    }
