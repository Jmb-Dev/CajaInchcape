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
using CajaIndigo.AppPersistencia.Class.MatrizDePago.Estructura;

namespace CajaIndigo.AppPersistencia.Class.MatrizDePago
{
    class MatrizDePago
    {
        public List<VIAS_PAGO> ObjDatosViasPago = new List<VIAS_PAGO>();
        public List<ViasPago> ViasPagoTransaccion = new List<ViasPago>();
        public string errormessage = "";
        ConexSAP connectorSap = new ConexSAP();


        public void viaspago(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_EXCEPCION, string P_CLASE_CUENTA, string P_HAS_CME, string P_LAND, string P_PROTESTO, List<ViasPago> P_CONDICIONES)
        {
            ObjDatosViasPago.Clear();
            ViasPagoTransaccion.Clear();
            errormessage = "";
            IRfcTable lt_VIAS_PAGOS;
           
           VIAS_PAGO VIAS_PAGOS_resp;

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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_MATRIZ_PAGO");
                BapiGetUser.SetValue("EXCEPCION", P_EXCEPCION);
                IRfcTable GralDat = BapiGetUser.GetTable("CONDICIONES");

                for (var i = 0; i < P_CONDICIONES.Count; i++)
                {
                    GralDat.Append();
                    GralDat.SetValue("ACC", P_CONDICIONES[i].Acc);
                    GralDat.SetValue("COND_PAGO", P_CONDICIONES[i].Cond_Pago);
                    GralDat.SetValue("CAJA", P_CONDICIONES[i].Caja);
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               }
                BapiGetUser.SetValue("CONDICIONES", GralDat);
                BapiGetUser.SetValue("CLASE_CUENTA", P_CLASE_CUENTA);
                BapiGetUser.SetValue("HAS_CME", P_HAS_CME);
                BapiGetUser.SetValue("LAND", P_LAND);
                BapiGetUser.SetValue("PROT", P_PROTESTO);
                BapiGetUser.Invoke(SapRfcDestination);
                lt_VIAS_PAGOS = BapiGetUser.GetTable("VIAS_PAGO");
                if (lt_VIAS_PAGOS.Count > 0)
                {
                    //LLenamos la tabla de salida lt_DatGen
                    for (int i = 0; i < lt_VIAS_PAGOS.RowCount; i++)
                    {
                        lt_VIAS_PAGOS.CurrentIndex = i;
                        VIAS_PAGOS_resp = new VIAS_PAGO();

                        VIAS_PAGOS_resp.VIA_PAGO = lt_VIAS_PAGOS[i].GetString("VIA_PAGO");
                        VIAS_PAGOS_resp.DESCRIPCION = lt_VIAS_PAGOS[i].GetString("DESCRIPCION");
                       ObjDatosViasPago.Add(VIAS_PAGOS_resp);
                    }
                }
                else
                {
                }
                GC.Collect();
            }
            else
            {              
                errormessage = retval;
                GC.Collect();
            }
        }
    }

}
