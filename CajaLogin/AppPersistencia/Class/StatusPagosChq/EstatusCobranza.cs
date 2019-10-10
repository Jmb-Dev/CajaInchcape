using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.StatusPagosChq.Estructura;


namespace CajaIndigo.AppPersistencia.Class.StatusPagosChq.Estructura
{
    class EstatusCobranza
    {

       
        public string errormessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public string protestado = "";
        ConexSAP connectorSap = new ConexSAP();
     
        public List<SE_ESTATUS> T_Retorno = new List<SE_ESTATUS>();

        public List<SE_ESTATUS> EstatusCobro(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_BUKRS, string P_KUNNR, string P_BSCHL, string P_UMSKZ, string P_UMSKS, string P_GJAHR)
        {
            try
            {
                T_Retorno.Clear();
                errormessage = "";
                status = "";
                message = "";
                stringRfc = "";
                protestado = "";
                //IRfcStructure ls_CIERRE_CAJA;
                //IRfcTable lt_CIERRE_CAJA;
                IRfcStructure lt_SE_STATUS;
                //IRfcTable lt_SE_STATUS;
                //CERR_CAJA CERR_CAJA_resp;
                SE_ESTATUS retorno; 
                //Conexion a SAP

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_STAT_COBRANZA");

                    BapiGetUser.SetValue("BUKRS", P_BUKRS);
                    BapiGetUser.SetValue("KUNNR", P_KUNNR);
                    BapiGetUser.SetValue("BSCHL", P_BSCHL);
                    BapiGetUser.SetValue("UMSKZ", P_UMSKZ);
                    BapiGetUser.SetValue("UMSKS", P_UMSKS);
                    BapiGetUser.SetValue("GJAHR", P_GJAHR);
                
                    BapiGetUser.Invoke(SapRfcDestination);
                    //LLenamos los datos que retorna la estructura de la RFC
                    //lt_CIERRE_CAJA = BapiGetUser.GetTable("ESTATUS");
                    protestado = BapiGetUser.GetString("PE_PROTESTADO");

                    lt_SE_STATUS = BapiGetUser.GetStructure("SE_ESTATUS");
                   // for (int i = 0; i < lt_SE_STATUS.Count(); i++)
                   // {
                       // lt_SE_STATUS.CurrentIndex = i;
                        retorno = new SE_ESTATUS();
                        retorno.TYPE = lt_SE_STATUS.GetString("TYPE");
                        retorno.ID = lt_SE_STATUS.GetString("ID");
                        retorno.NUMBER = lt_SE_STATUS.GetString("NUMBER");
                        retorno.MESSAGE = lt_SE_STATUS.GetString("MESSAGE");
                        retorno.LOG_NO = lt_SE_STATUS.GetString("LOG_NO");
                        retorno.LOG_MSG_NO = lt_SE_STATUS.GetString("LOG_MSG_NO");
                        retorno.MESSAGE_V1 = lt_SE_STATUS.GetString("MESSAGE_V1");
                        retorno.MESSAGE_V2 = lt_SE_STATUS.GetString("MESSAGE_V2");
                        retorno.MESSAGE_V3 = lt_SE_STATUS.GetString("MESSAGE_V3");
                        retorno.MESSAGE_V4 = lt_SE_STATUS.GetString("MESSAGE_V4");
                        retorno.PARAMETER = lt_SE_STATUS.GetString("PARAMETER");
                        retorno.ROW = lt_SE_STATUS.GetString("ROW");
                        retorno.FIELD = lt_SE_STATUS.GetString("FIELD");
                        retorno.SYSTEM = lt_SE_STATUS.GetString("SYSTEM");
                        T_Retorno.Add(retorno);
                  //  }

                }
                GC.Collect();
            }

            catch (InvalidCastException ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
            return T_Retorno;
        }
    }
}