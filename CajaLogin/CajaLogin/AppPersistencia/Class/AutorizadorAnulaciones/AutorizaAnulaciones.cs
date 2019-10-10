using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CajaIndu.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;
using CajaIndu.AppPersistencia.Class.Connections;
using SAP.Middleware.Connector;

namespace CajaIndu.AppPersistencia.Class.AutorizadorAnulaciones
{
    class AutorizaAnulaciones
    {
        public List<ESTADO> Retorno = new List<ESTADO>();
        public string Autorizado = "";
        public string errormessage = "";
        public bool Valido;
        ConexSAP connectorSap = new ConexSAP();


        public void anulacioncomprobantes(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_USUARIO, string P_PASSWORD2, string P_IDCAJA)
        {
            Retorno.Clear();
            Autorizado = "";
            errormessage = "";
            IRfcStructure lt_retorno;
            //IRfcTable lt_retorno;
            //string Autorizado;
            //bool Valido;

            ESTADO retorno_resp;
            

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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_CHECK_JEFE_CAJA");
                
                BapiGetUser.SetValue("PASSWORD", P_PASSWORD2);
                BapiGetUser.SetValue("ID_CAJA", P_IDCAJA);
                BapiGetUser.SetValue("USUARIO", P_USUARIO);

                BapiGetUser.Invoke(SapRfcDestination);

                Autorizado = BapiGetUser.GetString("VALIDO");
                if (Autorizado == "X")
                {
                    Valido = true;
                }

                //  lt_retorno = BapiGetUser.GetTable("ESTADO");
                lt_retorno = BapiGetUser.GetStructure("ESTADO");
                String Mensaje = "";
                if (lt_retorno.Count > 0)
                {
                    retorno_resp = new ESTADO();
                    for (int i = 0; i < lt_retorno.Count(); i++)
                    {
                       // lt_retorno.CurrentIndex = i;
                       
                        retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                        retorno_resp.ID = lt_retorno.GetString("ID");
                        retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                        retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                        retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                        retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                        retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                        retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                        if (i==0)
                        {
                            Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V1");
                        }
                        retorno_resp.MESSAGE_V2 = lt_retorno.GetString("MESSAGE_V2");
                        retorno_resp.MESSAGE_V3 = lt_retorno.GetString("MESSAGE_V3");
                        retorno_resp.MESSAGE_V4 = lt_retorno.GetString("MESSAGE_V4");
                        retorno_resp.PARAMETER = lt_retorno.GetString("PARAMETER");
                        retorno_resp.ROW = lt_retorno.GetString("ROW");
                        retorno_resp.FIELD = lt_retorno.GetString("FIELD");
                        retorno_resp.SYSTEM = lt_retorno.GetString("SYSTEM");
                        Retorno.Add(retorno_resp);
                    }
                    if (Mensaje != "")
                    {
                        System.Windows.MessageBox.Show(Mensaje);
                    }
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
