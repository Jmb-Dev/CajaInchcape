using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using CajaIndigo.AppPersistencia.Class.BusquedaAnulacion.Estructura;
using CajaIndigo.AppPersistencia.Class.AnulacionComprobantes.Estructura;
using CajaIndigo.AppPersistencia.Class.Connections;
using SAP.Middleware.Connector;


namespace CajaIndigo.AppPersistencia.Class.AnulacionComprobantes
{
    class AnulacionComprobantes
    {
        public string Efectivo = "";

        public List<RETORNO> Retorno = new List<RETORNO>();
        public List<RESUMEN_VP> resumenVp = new List<RESUMEN_VP>();
        public string errormessage = "";
        public string NumComprobante = "";
        public string Mensaje = "";
        ConexSAP connectorSap = new ConexSAP();


        public void anulacioncomprobantes(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_COMPROBANTE, string P_APROBADOR_ANULACION,
            string P_TXT_ANULACION, string P_USUARIO, string P_IDCAJA, string P_ACCION)
        {

            Retorno.Clear();
            errormessage = "";
            IRfcTable lt_retorno;

            RETORNO retorno_resp;

            //Conexion a SAP

            connectorSap.idioma = P_IDIOMA;
            connectorSap.idSistema = P_IDSISTEMA;
            connectorSap.instancia = P_INSTANCIA;
            connectorSap.mandante = P_MANDANTE;
            connectorSap.paswr = P_PASSWORD;
            connectorSap.sapRouter = P_SAPROUTER;
            connectorSap.user = P_UNAME;
            connectorSap.server = P_SERVER;

            string retval = connectorSap.connectionsSAP();
            //Si el valor de retorno es nulo o vacio, hay conexion a SAP y la RFC trae datos   
            if (string.IsNullOrEmpty(retval))
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_ANULA_REC_CAJA");
                BapiGetUser.SetValue("ID_COMPROBANTE", P_ID_COMPROBANTE);
                BapiGetUser.SetValue("APROBADOR_ANULACION", P_APROBADOR_ANULACION);
                BapiGetUser.SetValue("TXT_ANULACION", P_TXT_ANULACION);
                BapiGetUser.SetValue("ID_CAJA", P_IDCAJA);
                BapiGetUser.SetValue("USUARIO", P_USUARIO);
                BapiGetUser.SetValue("ACCION", P_ACCION);

        

                BapiGetUser.Invoke(SapRfcDestination);

                //lt_h_documentos = BapiGetUser.GetTable("CAB_COMP");
                //lt_d_documentos = BapiGetUser.GetTable("DET_COMP");
                lt_retorno = BapiGetUser.GetTable("RETORNO");

               
                if (lt_retorno.Count > 0)
                {
                    for (int i = 0; i < lt_retorno.Count(); i++)
                    {
                        lt_retorno.CurrentIndex = i;
                        retorno_resp = new RETORNO();
                        retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                        retorno_resp.ID = lt_retorno.GetString("ID");
                        retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                        retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                        retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                        retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                        retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                        retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                        if (lt_retorno.GetString("TYPE") == "S")
                        {
                            Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V4");
                            NumComprobante = lt_retorno.GetString("MESSAGE_V4");
                        }
                        if (lt_retorno.GetString("TYPE") == "E")
                        {
                            errormessage = errormessage + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V1");
                        } retorno_resp.MESSAGE_V2 = lt_retorno.GetString("MESSAGE_V2");
                        retorno_resp.MESSAGE_V3 = lt_retorno.GetString("MESSAGE_V3");
                        retorno_resp.MESSAGE_V4 = lt_retorno.GetString("MESSAGE_V4");
                        retorno_resp.PARAMETER = lt_retorno.GetString("PARAMETER");
                        retorno_resp.ROW = lt_retorno.GetString("ROW");
                        retorno_resp.FIELD = lt_retorno.GetString("FIELD");
                        retorno_resp.SYSTEM = lt_retorno.GetString("SYSTEM");
                        Retorno.Add(retorno_resp);
                    }
                }
            }
            else
            {
                errormessage = retval;
            }
            GC.Collect();
        }

        public void MontoEfectivo(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_USUARIO,
          string P_ID_APERTURA, string P_SOCIEDAD, string P_VIA_PAGO)
        {

            resumenVp.Clear();
            errormessage = "";
            IRfcTable lt_resumen_vp;

            RESUMEN_VP resumen_resp;


            //Conexion a SAP

            connectorSap.idioma = P_IDIOMA;
            connectorSap.idSistema = P_IDSISTEMA;
            connectorSap.instancia = P_INSTANCIA;
            connectorSap.mandante = P_MANDANTE;
            connectorSap.paswr = P_PASSWORD;
            connectorSap.sapRouter = P_SAPROUTER;
            connectorSap.user = P_UNAME;
            connectorSap.server = P_SERVER;

            string retval = connectorSap.connectionsSAP();
            //Si el valor de retorno es nulo o vacio, hay conexion a SAP y la RFC trae datos   
            if (string.IsNullOrEmpty(retval))
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_MON_EFEC");
                BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                BapiGetUser.SetValue("USUARIO", P_USUARIO);
                BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);
                BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                BapiGetUser.SetValue("VIA_PAGO", P_VIA_PAGO);

                BapiGetUser.Invoke(SapRfcDestination);

                lt_resumen_vp = BapiGetUser.GetTable("RESUMEN_VP");

                if (lt_resumen_vp.Count > 0)
                {
                    for (int i = 0; i < lt_resumen_vp.Count(); i++)
                    {
                        lt_resumen_vp.CurrentIndex = i;
                        resumen_resp = new RESUMEN_VP();
                        resumen_resp.LAND = lt_resumen_vp.GetString("LAND");
                        resumen_resp.ID_CAJA = lt_resumen_vp.GetString("ID_CAJA");
                        resumen_resp.SOCIEDAD = lt_resumen_vp.GetString("SOCIEDAD");
                        resumen_resp.SOCIEDAD_TXT = lt_resumen_vp.GetString("SOCIEDAD_TXT");
                        resumen_resp.VIA_PAGO = lt_resumen_vp.GetString("VIA_PAGO");
                        resumen_resp.TEXT1 = lt_resumen_vp.GetString("TEXT1");
                        resumen_resp.MONEDA = lt_resumen_vp.GetString("MONEDA");
                        resumen_resp.MONTO = lt_resumen_vp.GetString("MONTO");
                        Efectivo = Convert.ToString(Convert.ToDouble(resumen_resp.MONTO));
                        
                        resumen_resp.CANT_DOCS = lt_resumen_vp.GetString("CANT_DOCS");
                        resumenVp.Add(resumen_resp);
                    }

                    Double Efect = Convert.ToDouble(Efectivo);
                    CultureInfo elGR = CultureInfo.CreateSpecificCulture("el-GR");
                    Console.WriteLine(Efect.ToString("0,0", elGR));
                    Console.WriteLine(String.Format(elGR, "{0:0,0}", Efect));
                    Efectivo = Efect.ToString("0,0", elGR);
                    Efectivo = String.Format(elGR, "{0:0,0}", Efect);
                }
            }
            else
            {
                errormessage = retval;
            }
            GC.Collect();
        }
    }
}
