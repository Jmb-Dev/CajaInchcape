using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.MaestroTarjetas;
using CajaIndu.AppPersistencia.Class.MaestroTarjetas.Estructura;

namespace CajaIndu.AppPersistencia.Class.MaestroTarjetas
{
    class MaestroTarjetas
    {
    
        public string pagomessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public int id_error = 0;
        
        ConexSAP connectorSap = new ConexSAP();
        public List<LISTATARJETAS> T_Retorno = new List<LISTATARJETAS>();


        public void maestrotarjetas(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_PAIS, string P_VPAGO) 
        {
            try
            {
                T_Retorno.Clear();
                
                IRfcTable lt_TARJETAS;
               
                LISTATARJETAS retorno;
                
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_TARJETAS");

                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("VPAGO", P_VPAGO.Substring(0,1));
                    
                    BapiGetUser.Invoke(SapRfcDestination);
                    //LLenamos los datos que retorna la estructura de la RFC
                    //pagomessage = BapiGetUser.GetString("E_MSJ");
                    //id_error = BapiGetUser.GetInt("E_ID_MSJ");
                    // message = BapiGetUser.GetString("E_AUGBL");

                    lt_TARJETAS = BapiGetUser.GetTable("LISTATARJETAS");
                    
                    for (int i = 0; i < lt_TARJETAS.Count(); i++)
                    {
                        lt_TARJETAS.CurrentIndex = i;
                        retorno = new LISTATARJETAS();
                        retorno.MANDT = lt_TARJETAS.GetString("MANDT");
                        retorno.LAND = lt_TARJETAS.GetString("LAND");
                        retorno.VIA_PAGO = lt_TARJETAS.GetString("VIA_PAGO");
                        retorno.CCINS = lt_TARJETAS.GetString("CCINS");
                        retorno.VTEXT = lt_TARJETAS.GetString("VTEXT");
                        T_Retorno.Add(retorno);
                    }
                    
                    }
                GC.Collect();
                }
           

            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                GC.Collect();
            }
            // return T_Retorno;
        }
    }
}



