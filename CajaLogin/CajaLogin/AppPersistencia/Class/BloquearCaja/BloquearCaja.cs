using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura;
using CajaIndu.AppPersistencia.Class.UsuariosCaja;


namespace CajaIndu.AppPersistencia.Class.BloquearCaja
{
    class BloquearCaja
    {
        ConexSAP connectorSap = new ConexSAP();
        string errormessage = "";


        public void bloqueardesbloquearcaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, List<LOG_APERTURA> P_LOGAPERTURA)
        {

            //IRfcTable lt_t_documentos;
            //IRfcStructure lt_retorno;

            //  PART_ABIERTAS  PART_ABIERTAS_resp;
           

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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_BLOQ_DES_CAJA");
                //BapiGetUser.SetValue("LOG_APERTURA", P_LOGAPERTURA);
                IRfcStructure GralDat = BapiGetUser.GetStructure("LOG_APERTURA");

                for (var i = 0; i < P_LOGAPERTURA.Count; i++)
                {
                    //GralDat.Append();
                    GralDat.SetValue("MANDT", P_LOGAPERTURA[i].MANDT);
                    GralDat.SetValue("ID_REGISTRO", P_LOGAPERTURA[i].ID_REGISTRO);
                    GralDat.SetValue("LAND", P_LOGAPERTURA[i].LAND);
                    GralDat.SetValue("ID_CAJA", P_LOGAPERTURA[i].ID_CAJA);
                    GralDat.SetValue("USUARIO", P_LOGAPERTURA[i].USUARIO);
                    GralDat.SetValue("FECHA", P_LOGAPERTURA[i].FECHA);
                    GralDat.SetValue("HORA", P_LOGAPERTURA[i].HORA);
                    GralDat.SetValue("MONTO", P_LOGAPERTURA[i].MONTO);
                    GralDat.SetValue("MONEDA", P_LOGAPERTURA[i].MONEDA);
                    GralDat.SetValue("TIPO_REGISTRO", P_LOGAPERTURA[i].TIPO_REGISTRO);
                    GralDat.SetValue("ID_APERTURA", P_LOGAPERTURA[i].ID_APERTURA);
                    GralDat.SetValue("TXT_CIERRE", P_LOGAPERTURA[i].TXT_CIERRE);
                    GralDat.SetValue("BLOQUEO", P_LOGAPERTURA[i].BLOQUEO);
                    GralDat.SetValue("USUARIO_BLOQ", P_LOGAPERTURA[i].USUARIO_BLOQ);

                }
                BapiGetUser.SetValue("LOG_APERTURA", GralDat);
                BapiGetUser.Invoke(SapRfcDestination);
                

                //lt_t_documentos = BapiGetUser.GetTable("T_DOCUMENTOS");
                //lt_retorno = BapiGetUser.GetStructure("SE_ESTATUS");
                //lt_PART_ABIERTAS = BapiGetUser.GetTable("ZCLSP_TT_LISTA_DOCUMENTOS");
                try
                {
                    

                    String Mensaje = "";
                    //if (lt_retorno.Count > 0)
                    //{
                    //    retorno_resp = new ESTADO();
                    //    for (int i = 0; i < lt_retorno.Count(); i++)
                    //    {
                    //        // lt_retorno.CurrentIndex = i;

                    //        retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                    //        retorno_resp.ID = lt_retorno.GetString("ID");
                    //        retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                    //        retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                    //        retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                    //        retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                    //        retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                    //        retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                    //        if (i == 0)
                    //        {
                    //            Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V1");
                    //        }
                    //        retorno_resp.MESSAGE_V2 = lt_retorno.GetString("MESSAGE_V2");
                    //        retorno_resp.MESSAGE_V3 = lt_retorno.GetString("MESSAGE_V3");
                    //        retorno_resp.MESSAGE_V4 = lt_retorno.GetString("MESSAGE_V4");
                    //        retorno_resp.PARAMETER = lt_retorno.GetString("PARAMETER");
                    //        retorno_resp.ROW = lt_retorno.GetString("ROW");
                    //        retorno_resp.FIELD = lt_retorno.GetString("FIELD");
                    //        retorno_resp.SYSTEM = lt_retorno.GetString("SYSTEM");
                    //        Retorno.Add(retorno_resp);
                    //    }
                    //    System.Windows.MessageBox.Show(Mensaje);
                    //}
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

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
