using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using SAP.Middleware.Connector;
using System.Configuration;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.AperturaCaja.Estructura;

namespace CajaIndigo.AppPersistencia.Class.AperturaCaja
{
    class AperturaCaja
    {

        //public List<USR_CAJA> ObjDatosLogin = new List<USR_CAJA>();
        public string errormessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        ConexSAP connectorSap = new ConexSAP();

        public void aperturacaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER
            , string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_MONTO, string P_PAIS, string P_MONEDA, string P_TIPO_REGISTRO, string P_EQUIPO)
        {
            try
            {
                errormessage = "";
                status = "";
                message = "";
                stringRfc = "";
                IRfcStructure ls_APER_CAJA;
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_GF_RFC_APER_CAJA");

                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    if (P_MONTO == "")
                    {
                        P_MONTO = "0";
                    }
                    BapiGetUser.SetValue("MONTO", P_MONTO);

                    BapiGetUser.SetValue("MONEDA", P_MONEDA);
                    BapiGetUser.SetValue("TIPO_REGISTRO", P_TIPO_REGISTRO);
                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("USUARIO", P_UNAME);
                    BapiGetUser.SetValue("EQUIPO", P_EQUIPO);

                    BapiGetUser.Invoke(SapRfcDestination);
                    ls_APER_CAJA = BapiGetUser.GetStructure("ESTATUS");

                    stringRfc = Convert.ToString(BapiGetUser.GetValue("ESTATUS"));
                    //LLenamos los datos que retorna la estructura de la RFC
                    status = ls_APER_CAJA.GetString("TYPE");
                    message = ls_APER_CAJA.GetString("MESSAGE");


                }
                else
                ////Si el valor de retorno es distinto nulo o vacio,se emite el mensaje de error.
                {
                    errormessage = retval;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
            GC.Collect();
        }
    }
     
}
