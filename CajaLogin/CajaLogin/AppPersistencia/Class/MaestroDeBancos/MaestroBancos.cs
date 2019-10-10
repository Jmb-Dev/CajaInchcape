using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Shapes;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.MaestroDeBancos.Estructura;

namespace CajaIndu.AppPersistencia.Class.MaestroDeBancos
{
    class MaestroBancos
    {
        public string pagomessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public int id_error = 0;
        ConexSAP connectorSap = new ConexSAP();
        public List<LISTABANCOS> T_Retorno = new List<LISTABANCOS>();
        public List<BANCOS_PROPIOS> T_Retorno2 = new List<BANCOS_PROPIOS>();

        public void maestrobancos(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_PAIS, string P_MONEDA, string P_SOCIEDAD) 
        {
            try
            {
                T_Retorno.Clear();
                T_Retorno2.Clear();
                IRfcTable lt_BANCOS;
                IRfcTable lt_BANCOS2;
                LISTABANCOS retorno;
                BANCOS_PROPIOS retorno2;
                //Conexion a SAP
                T_Retorno.Clear();
                T_Retorno2.Clear();

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_BANCO");

                    BapiGetUser.SetValue("LAND1", P_PAIS);
                    BapiGetUser.SetValue("PAY_CURRENCY", P_MONEDA);
                    BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);


                    BapiGetUser.Invoke(SapRfcDestination);
                    //LLenamos los datos que retorna la estructura de la RFC
                    //pagomessage = BapiGetUser.GetString("E_MSJ");
                    //id_error = BapiGetUser.GetInt("E_ID_MSJ");
                    // message = BapiGetUser.GetString("E_AUGBL");

                    lt_BANCOS = BapiGetUser.GetTable("LISTABANCOS");
                    lt_BANCOS2 = BapiGetUser.GetTable("BANCOS_PROPIOS");
                    for (int i = 0; i < lt_BANCOS.Count(); i++)
                    {
                        lt_BANCOS.CurrentIndex = i;
                        retorno = new LISTABANCOS();
                        retorno.BANKL = lt_BANCOS.GetString("BANKL");
                        retorno.BANKA = lt_BANCOS.GetString("BANKA");

                        T_Retorno.Add(retorno);
                    }
                    for (int i = 0; i < lt_BANCOS2.Count(); i++)
                    {
                        lt_BANCOS2.CurrentIndex = i;
                        retorno2 = new BANCOS_PROPIOS();
                        retorno2.BUKRS = lt_BANCOS2.GetString("BUKRS");
                        retorno2.HBKID = lt_BANCOS2.GetString("HBKID");
                        retorno2.HKTID = lt_BANCOS2.GetString("HKTID");
                        retorno2.BANKN = lt_BANCOS2.GetString("BANKN");
                        retorno2.BANKL = lt_BANCOS2.GetString("BANKL");
                        retorno2.BANKA = lt_BANCOS2.GetString("BANKA");
                        retorno2.WAERS = lt_BANCOS2.GetString("WAERS");
                        retorno2.TEXT1 = lt_BANCOS2.GetString("TEXT1");

                        T_Retorno2.Add(retorno2);
                    }

                }
            }

            catch (InvalidCastException ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
            // return T_Retorno;
        }
    }
}



