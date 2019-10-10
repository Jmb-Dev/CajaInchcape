using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using SAP.Middleware.Connector;
using System.Configuration;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.Login.Estructura;


namespace CajaIndu.AppPersistencia.Class.Login.Estructura
{
    class LoginSAP 
    {
       public List<USR_CAJA> ObjDatosLogin = new List<USR_CAJA>();
       public string errormessage = "";
       ConexSAP connectorSap = new ConexSAP();

       public void datoslogin(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER
           , string P_SERVER, string P_IDIOMA, string P_TEMPORAL,string P_EQUIPO)  
        {
            ObjDatosLogin.Clear();
            IRfcTable lt_USR_CAJA;  

            USR_CAJA  USR_CAJA_resp;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_USR_CAJA");
                    BapiGetUser.SetValue("UNAME", P_UNAME);
                    BapiGetUser.SetValue("TEMPORAL", P_TEMPORAL);
                    BapiGetUser.SetValue("EQUIPO", P_EQUIPO);
                    try
                    {
                        BapiGetUser.Invoke(SapRfcDestination);
                        lt_USR_CAJA = BapiGetUser.GetTable("USR_CAJA");

                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_USR_CAJA.RowCount; i++)
                        {
                            lt_USR_CAJA.CurrentIndex = i;
                            USR_CAJA_resp = new USR_CAJA();

                            USR_CAJA_resp.ID_CAJA = lt_USR_CAJA[i].GetString("ID_CAJA");
                            USR_CAJA_resp.NOM_CAJA = lt_USR_CAJA[i].GetString("NOM_CAJA");
                            USR_CAJA_resp.USUARIO = lt_USR_CAJA[i].GetString("USUARIO");
                            USR_CAJA_resp.TIPO_USUARIO = lt_USR_CAJA[i].GetString("TIPO_USUARIO");
                            USR_CAJA_resp.LAND = lt_USR_CAJA[i].GetString("LAND");
                            USR_CAJA_resp.MONEDA = lt_USR_CAJA[i].GetString("WAERS");

                            ObjDatosLogin.Add(USR_CAJA_resp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                        errormessage = ex.Message;
                    }
                    if (errormessage == "NO_DATA")
                    {
                        errormessage = "No existe usuario registrado en los datos maestros‏";
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

