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
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.Monitor.Estructura;
using CajaIndu.AppPersistencia.Class.PartidasAbiertas.Estructura;

namespace CajaIndu.AppPersistencia.Class.Monitor.Estructura
{
    class Monitor
    {
        public List<T_DOCUMENTOS> ObjDatosMonitor = new List<T_DOCUMENTOS>();
        public string errormessage = "";
        ConexSAP connectorSap = new ConexSAP();

        public void monitor(string P_DATUM, string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_SOCIEDAD)
        {
            try
            {
                ObjDatosMonitor.Clear();
                        IRfcTable lt_GET_MONITOR;

                        T_DOCUMENTOS GET_MONITOR_resp;

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

                            IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_MONITOR");
                            BapiGetUser.SetValue("DATUM",Convert.ToDateTime( P_DATUM));
                            BapiGetUser.SetValue("I_BUKRS", P_SOCIEDAD);

                            BapiGetUser.Invoke(SapRfcDestination);

                            lt_GET_MONITOR = BapiGetUser.GetTable("T_DOCUMENTOS");
                            //lt_PART_ABIERTAS = BapiGetUser.GetTable("ZCLSP_TT_LISTA_DOCUMENTOS");
                            if (lt_GET_MONITOR.Count > 0)
                            {
                                //LLenamos la tabla de salida lt_DatGen
                                for (int i = 0; i < lt_GET_MONITOR.RowCount; i++)
                                {
                                    lt_GET_MONITOR.CurrentIndex = i;
                                    GET_MONITOR_resp = new T_DOCUMENTOS();

                                    GET_MONITOR_resp.SOCIEDAD = lt_GET_MONITOR[i].GetString("SOCIEDAD");
                                    GET_MONITOR_resp.NDOCTO = lt_GET_MONITOR[i].GetString("NDOCTO");
                                    GET_MONITOR_resp.NREF = lt_GET_MONITOR[i].GetString("NREF");
                                    GET_MONITOR_resp.CLASE_CUENTA = lt_GET_MONITOR[i].GetString("CLASE_CUENTA");
                                    GET_MONITOR_resp.CLASE_DOC = lt_GET_MONITOR[i].GetString("CLASE_DOC");
                                    GET_MONITOR_resp.COD_CLIENTE = lt_GET_MONITOR[i].GetString("COD_CLIENTE");
                                    GET_MONITOR_resp.RUTCLI = lt_GET_MONITOR[i].GetString("RUTCLI");
                                    GET_MONITOR_resp.NOMCLI = lt_GET_MONITOR[i].GetString("NOMCLI");
                                    GET_MONITOR_resp.CEBE = lt_GET_MONITOR[i].GetString("CEBE");
                                    GET_MONITOR_resp.FECHA_DOC = lt_GET_MONITOR[i].GetString("FECHA_DOC");
                                    GET_MONITOR_resp.FECVENCI = lt_GET_MONITOR[i].GetString("FECVENCI");
                                    GET_MONITOR_resp.DIAS_ATRASO = lt_GET_MONITOR[i].GetString("DIAS_ATRASO");
                                    GET_MONITOR_resp.ESTADO = lt_GET_MONITOR[i].GetString("ESTADO");
                                    GET_MONITOR_resp.ICONO = lt_GET_MONITOR[i].GetString("ICONO");
                                    GET_MONITOR_resp.MONEDA = lt_GET_MONITOR[i].GetString("MONEDA");
                                    GET_MONITOR_resp.ACC = lt_GET_MONITOR[i].GetString("ACC");
                                    GET_MONITOR_resp.CLASE_CUENTA = lt_GET_MONITOR[i].GetString("CLASE_CUENTA");
                                    GET_MONITOR_resp.COND_PAGO = lt_GET_MONITOR[i].GetString("COND_PAGO");
                                    GET_MONITOR_resp.CME = lt_GET_MONITOR[i].GetString("CME");
                                    GET_MONITOR_resp.CONTROL_CREDITO = lt_GET_MONITOR[i].GetString("CONTROL_CREDITO");
                                    string Monto = "";
                                    int indice = 0;
                                    if (lt_GET_MONITOR[i].GetString("MONTO").Contains(","))
                                    {
                                       indice = lt_GET_MONITOR[i].GetString("MONTO").IndexOf(',');
                                       Monto = lt_GET_MONITOR[i].GetString("MONTO").Substring(0, indice - 1);
                                       GET_MONITOR_resp.MONTOF = Monto;
                                    }
                                    else
                                    {
                                         GET_MONITOR_resp.MONTOF = lt_GET_MONITOR[i].GetString("MONTOF");
                                    }
                                    if (lt_GET_MONITOR[i].GetString("MONTO").Contains(","))
                                    {
                                        indice = lt_GET_MONITOR[i].GetString("MONTO").IndexOf(',');
                                        Monto = lt_GET_MONITOR[i].GetString("MONTO").Substring(0, indice - 1);
                                        GET_MONITOR_resp.MONTO = Monto;
                                    }
                                    else
                                    {
                                        GET_MONITOR_resp.MONTO = lt_GET_MONITOR[i].GetString("MONTO");
                                    }
                                    if (lt_GET_MONITOR[i].GetString("MONTO_ABONADO").Contains(","))
                                    {
                                        indice = lt_GET_MONITOR[i].GetString("MONTO_ABONADO").IndexOf(',');
                                        Monto = lt_GET_MONITOR[i].GetString("MONTO_ABONADO").Substring(0, indice - 1);
                                       GET_MONITOR_resp.MONTOF = Monto;
                                    }
                                    else
                                    {                                 
                                        GET_MONITOR_resp.MONTOF_ABON = lt_GET_MONITOR[i].GetString("MONTOF_ABON");
                                    }
                                     if (lt_GET_MONITOR[i].GetString("MONTO_PAGAR").Contains(","))
                                    {
                                       indice = lt_GET_MONITOR[i].GetString("MONTO_PAGAR").IndexOf(',');
                                       Monto = lt_GET_MONITOR[i].GetString("MONTO_PAGAR").Substring(0,indice-1);
                                       GET_MONITOR_resp.MONTOF = Monto;
                                    }
                                    else
                                    {  
                                    GET_MONITOR_resp.MONTOF_PAGAR = lt_GET_MONITOR[i].GetString("MONTOF_PAGAR");
                                     }

                                    ObjDatosMonitor.Add(GET_MONITOR_resp);


                                }
                            }
                            //else
                            //{
                            //    MessageBox.Show("No existen registros para este fecha");
                            //}
                            GC.Collect();
                        }
                        else
                        {
                            errormessage = retval;
                        }
                        GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
       
    }
}

