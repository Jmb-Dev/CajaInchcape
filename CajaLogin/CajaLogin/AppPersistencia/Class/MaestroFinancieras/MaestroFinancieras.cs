using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.MaestroFinancieras.Estructura;

namespace CajaIndu.AppPersistencia.Class.MaestroFinancieras
{
    class MaestroFinancieras
    {

        public string pagomessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public int id_error = 0;
        ConexSAP connectorSap = new ConexSAP();
        public List<LIST_CARTA> T_Retorno = new List<LIST_CARTA>();


        public void maestroifinan(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_PAIS, string P_VPAGO, string P_SOCIEDAD) 
        {
            try
            {
                T_Retorno.Clear();
              
                IRfcTable lt_BANCOS;
               
                LIST_CARTA retorno;
               
                //Conexion a SAP
                T_Retorno.Clear();
                

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_CARTA_COURSE");

                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("VPAGO", P_VPAGO);
                    BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);


                    BapiGetUser.Invoke(SapRfcDestination);
                    //LLenamos los datos que retorna la estructura de la RFC
                    //pagomessage = BapiGetUser.GetString("E_MSJ");
                    //id_error = BapiGetUser.GetInt("E_ID_MSJ");
                    // message = BapiGetUser.GetString("E_AUGBL");

                    lt_BANCOS = BapiGetUser.GetTable("LIST_CARTA");
                   
                    for (int i = 0; i < lt_BANCOS.Count(); i++)
                    {
                        lt_BANCOS.CurrentIndex = i;
                        retorno = new LIST_CARTA();
                        retorno.MANDT = lt_BANCOS.GetString("MANDT");
                        retorno.LAND = lt_BANCOS.GetString("LAND");
                        retorno.BUKRS = lt_BANCOS.GetString("BUKRS");
                        retorno.CODIN = lt_BANCOS.GetString("CODIN");
                        retorno.KUNNR = lt_BANCOS.GetString("KUNNR");
                        retorno.MCOD1 = lt_BANCOS.GetString("MCOD1");
                        T_Retorno.Add(retorno);
                    }
                   

                }
            }

            catch (InvalidCastException ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                //System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
            // return T_Retorno;
        }
    }
}



