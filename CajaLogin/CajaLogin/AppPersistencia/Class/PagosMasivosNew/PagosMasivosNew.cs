using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.CierreCaja.Estructura;

namespace CajaIndu.AppPersistencia.Class.PagosMasivosNew
{
    class PagosMasivosNew 
    {

        public List<ESTATUS> objReturn2 = new List<ESTATUS>();
        public string message = "";
        public string errormessage = "";
        ConexSAP connectorSap = new ConexSAP();

        public void pagosmasivos(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA
            , string P_LAND, string P_FECHA, string P_FILE, string P_ID_APERTURA, string P_ID_CAJA, string P_PAY_CURRENCY, List<PagosMasivosNuevo> ListaExc)
        {

            objReturn2.Clear();
            errormessage = "";
            message = "";
            try
            {
                ESTATUS p_return;
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_PAGO_MASIVO");
                    BapiGetUser.SetValue("LAND", P_LAND);
                    BapiGetUser.SetValue("FECHA", Convert.ToDateTime(P_FECHA.Substring(0,10)));
                    //BapiGetUser.SetValue("FILEN", P_FILE);
                    BapiGetUser.SetValue("PAY_CURRENCY", P_PAY_CURRENCY);
                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);

                    IRfcTable GralDat = BapiGetUser.GetTable("T_EXCEL");

                    for (var i = 0; i < ListaExc.Count; i++)
                    {
                        GralDat.Append();
                        GralDat.SetValue("ROW", ListaExc[i].ROW);
                        GralDat.SetValue("COL", ListaExc[i].COL);
                        GralDat.SetValue("VALUE", ListaExc[i].VALUE);
                    }
                    BapiGetUser.SetValue("T_EXCEL", GralDat);


                    BapiGetUser.Invoke(SapRfcDestination);

                    IRfcTable retorno = BapiGetUser.GetTable("ESTATUS");

                    for (var i = 0; i < retorno.RowCount; i++)
                    {
                        retorno.CurrentIndex = i;

                        p_return = new ESTATUS();

                        p_return.TYPE = retorno[i].GetString("TYPE");
                        if (retorno.GetString("TYPE") == "S")
                        {
                            message = message + " - " + retorno[i].GetString("MESSAGE");
                        }
                        if (retorno.GetString("TYPE") == "E")
                        {
                            errormessage = errormessage + " - " + retorno[i].GetString("MESSAGE");
                        }
                        p_return.ID = retorno[i].GetString("ID");
                        p_return.NUMBER = retorno[i].GetString("NUMBER");
                        p_return.MESSAGE = retorno[i].GetString("MESSAGE");
                        p_return.LOG_NO = retorno[i].GetString("LOG_NO");
                        p_return.LOG_MSG_NO = retorno[i].GetString("LOG_MSG_NO");
                        p_return.MESSAGE_V1 = retorno[i].GetString("MESSAGE_V1");
                        p_return.MESSAGE_V2 = retorno[i].GetString("MESSAGE_V2");
                        p_return.MESSAGE_V3 = retorno[i].GetString("MESSAGE_V3");
                        p_return.MESSAGE_V4 = retorno[i].GetString("MESSAGE_V4");
                        p_return.PARAMETER = retorno[i].GetString("PARAMETER");
                        p_return.ROW = retorno[i].GetString("ROW");
                        p_return.FIELD = retorno[i].GetString("FIELD");
                        p_return.SYSTEM = retorno[i].GetString("SYSTEM");
                        objReturn2.Add(p_return);

                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.Forms.MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }
    }
}