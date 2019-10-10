using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SAP.Middleware.Connector;

namespace CajaIndu.AppPersistencia.Class.Connections
{
    class ConexSAP
    {
        public string server { get; set; }
        public string instancia { get; set; }
        public string idSistema { get; set; }
        public string sapRouter { get; set; }
        public string mandante { get; set; }
        public string user { get; set; }
        public string paswr { get; set; }
        public string idioma { get; set; }
        public RfcConfigParameters connectorConfig { get; set; }

        public string ToJson()
        {
            System.Web.Script.Serialization.JavaScriptSerializer jsonSerializer =
                        new System.Web.Script.Serialization.JavaScriptSerializer();

            return jsonSerializer.Serialize(this);
        }

        //Datos de conexión Indumotora DBM
        public RfcConfigParameters SAPConector()
        {
            RfcConfigParameters SapConnector = new RfcConfigParameters();

            SapConnector.Add(RfcConfigParameters.Name, idSistema);
            SapConnector.Add(RfcConfigParameters.AppServerHost, server);
            SapConnector.Add(RfcConfigParameters.SAPRouter, sapRouter);
            SapConnector.Add(RfcConfigParameters.SystemNumber, instancia);
            SapConnector.Add(RfcConfigParameters.User, user);
            SapConnector.Add(RfcConfigParameters.Password, paswr);
            SapConnector.Add(RfcConfigParameters.Client, mandante);
            SapConnector.Add(RfcConfigParameters.Language, "ES");
            SapConnector.Add(RfcConfigParameters.PoolSize, "10");
            SapConnector.Add(RfcConfigParameters.IdleTimeout, "10");

            return SapConnector;
        }


        public string connectionsSAP()
        {
            string mensaje = null;
            connectorConfig = SAPConector();

            try
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorConfig);
                SapRfcDestination.Ping();
            }
            catch (RfcLogonException ex)
            {
                mensaje = ex.Message;
            }
            catch (RfcCommunicationException exp)
            {
                mensaje = exp.Message;
            }
            finally { }

            return mensaje;
        }
    }
}
