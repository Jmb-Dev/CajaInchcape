using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.RendicionCaja;
using CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.ArqueoCaja.Estructura;

namespace CajaIndigo.AppPersistencia.Class.ArqueoCaja
{
    class ArqueoCaja
    {
        public string errormessage = "";
        public string message = "";
        public string diferencia = "";
        public string id_arqueo = "";
        //DETALLE_VP
        public List<DETALLE_VP> detalle_viapago = new List<DETALLE_VP>(); 
        //RESUMEN_VP
        public List<RESUMEN_VP> resumen_viapago = new List<RESUMEN_VP>();
        //DETALLE_REND
        public List<DETALLE_ARQUEO> detalle_rend = new List<DETALLE_ARQUEO>();
        ConexSAP connectorSap = new ConexSAP();
        public List<ESTATUS> T_Retorno = new List<ESTATUS>();

        public void arqueocaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER
            , string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_DATUMDESDE, string P_DATUMHASTA, string P_USUARIO
            , string P_PAIS, string P_MONEDALOCAL, string P_ID_APERTURA,  string P_ID_CIERRE, string P_IND_ARQUEO,string P_ID_ARQUEO_IN
            , string P_MTO_APERTURA, List<DETALLE_ARQUEO> P_TOTALEFECTIVO)
        {
            try
            {
                ESTATUS retorno;
                DETALLE_VP detallevp;
                RESUMEN_VP resumenvp;
                DETALLE_ARQUEO detallerend;
                T_Retorno.Clear();
                detalle_rend.Clear();
                detalle_viapago.Clear();
                resumen_viapago.Clear();
                errormessage = "";
                message = "";
                diferencia = "";
                id_arqueo = "";

                IRfcTable ls_RETORNO;
                IRfcTable lt_DETALLE_VP;
                IRfcTable lt_RESUMEN_VP;
                IRfcTable lt_DETALLE_REND;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_ARQUEO_CAJA_2");
                    
                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_USUARIO);
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA); // Buscar en log de apertura
                    BapiGetUser.SetValue("MONEDA_LOCAL", P_MONEDALOCAL); // Moneda
                    BapiGetUser.SetValue("ID_ARQUEO_IN", P_ID_ARQUEO_IN);// ""
                    BapiGetUser.SetValue("ID_CIERRE", P_ID_CIERRE); // ""
                    BapiGetUser.SetValue("IND_ARQUEO", P_IND_ARQUEO); //"A" Arqueo
                    BapiGetUser.SetValue("MTO_APERTURA", P_MTO_APERTURA); // Buscar en log de apertura

                    IRfcTable GralDat2 = BapiGetUser.GetTable("DETALLE_ARQUEO");
                    try
                    {
                        for (var i = 0; i < P_TOTALEFECTIVO.Count; i++)
                        {
                            GralDat2.Append();
                            GralDat2.SetValue("LAND", P_TOTALEFECTIVO[i].LAND);
                            GralDat2.SetValue("ID_CAJA", P_TOTALEFECTIVO[i].ID_CAJA);
                            GralDat2.SetValue("USUARIO", P_TOTALEFECTIVO[i].USUARIO);
                            GralDat2.SetValue("SOCIEDAD", P_TOTALEFECTIVO[i].SOCIEDAD);
                            GralDat2.SetValue("FECHA_REND",Convert.ToDateTime( P_TOTALEFECTIVO[i].FECHA_REND));
                            GralDat2.SetValue("HORA_REND", Convert.ToDateTime(P_TOTALEFECTIVO[i].HORA_REND));
                            GralDat2.SetValue("MONEDA", P_TOTALEFECTIVO[i].MONEDA);
                            GralDat2.SetValue("VIA_PAGO", P_TOTALEFECTIVO[i].VIA_PAGO);
                            GralDat2.SetValue("TIPO_MONEDA", P_TOTALEFECTIVO[i].TIPO_MONEDA);
                            GralDat2.SetValue("CANTIDAD_MON", P_TOTALEFECTIVO[i].CANTIDAD_MON);
                            GralDat2.SetValue("SUMA_MON_BILL", P_TOTALEFECTIVO[i].SUMA_MON_BILL);
                            GralDat2.SetValue("CANTIDAD_DOC", P_TOTALEFECTIVO[i].CANTIDAD_DOC);
                            GralDat2.SetValue("SUMA_DOCS", P_TOTALEFECTIVO[i].SUMA_DOCS);
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                        //System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                    BapiGetUser.SetValue("DETALLE_ARQUEO", GralDat2);
                    

                    BapiGetUser.Invoke(SapRfcDestination);

                    diferencia = BapiGetUser.GetString("DIFERENCIA");
                    if (diferencia.Contains(","))
                    {
                        diferencia = diferencia.Replace(",", "");
                        diferencia = diferencia.Substring(0, diferencia.Length - 2);
                    }
                    id_arqueo = BapiGetUser.GetString("ID_ARQUEO");
                    
                    
                    lt_DETALLE_VP = BapiGetUser.GetTable("DETALLE_VP"); //

                    try
                    {
                        for (int i = 0; i < lt_DETALLE_VP.Count(); i++)
                        {
                            lt_DETALLE_VP.CurrentIndex = i;
                            detallevp = new DETALLE_VP();
                            detallevp.SOCIEDAD = lt_DETALLE_VP.GetString("SOCIEDAD");
                            detallevp.SOCIEDAD_TXT = lt_DETALLE_VP.GetString("SOCIEDAD_TXT");
                            detallevp.ID_COMPROBANTE = lt_DETALLE_VP.GetString("ID_COMPROBANTE");
                            detallevp.ID_DETALLE = lt_DETALLE_VP.GetString("ID_DETALLE");
                            detallevp.VIA_PAGO = lt_DETALLE_VP.GetString("VIA_PAGO");
                            detallevp.MONTO = lt_DETALLE_VP.GetString("MONTO");
                            detallevp.MONEDA = lt_DETALLE_VP.GetString("MONEDA");
                            detallevp.BANCO = lt_DETALLE_VP.GetString("BANCO");
                            detallevp.BANCO_TXT = lt_DETALLE_VP.GetString("BANCO_TXT");
                            detallevp.EMISOR = lt_DETALLE_VP.GetString("EMISOR");
                            detallevp.NUM_CHEQUE = lt_DETALLE_VP.GetString("NUM_CHEQUE");
                            detallevp.COD_AUTORIZACION = lt_DETALLE_VP.GetString("COD_AUTORIZACION");
                            detallevp.CLIENTE = lt_DETALLE_VP.GetString("CLIENTE");
                            detallevp.NRO_DOCUMENTO = lt_DETALLE_VP.GetString("NRO_DOCUMENTO");
                            detallevp.NUM_CUOTAS = lt_DETALLE_VP.GetString("NUM_CUOTAS");
                            detallevp.FECHA_VENC = lt_DETALLE_VP.GetString("FECHA_VENC");
                            detallevp.FECHA_EMISION = lt_DETALLE_VP.GetString("FECHA_EMISION");
                            detallevp.NOTA_VENTA = lt_DETALLE_VP.GetString("NOTA_VENTA");
                            detallevp.TEXTO_POSICION = lt_DETALLE_VP.GetString("TEXTO_POSICION");
                            detallevp.NULO = lt_DETALLE_VP.GetString("NULO");
                            detallevp.NRO_REFERENCIA = lt_DETALLE_VP.GetString("NRO_REFERENCIA");
                            detallevp.TIPO_DOCUMENTO = lt_DETALLE_VP.GetString("TIPO_DOCUMENTO");
                            detalle_viapago.Add(detallevp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_RESUMEN_VP = BapiGetUser.GetTable("RESUMEN_VP"); //Resumen de Vias de Pago
                    try
                    {
                        for (int i = 0; i < lt_RESUMEN_VP.Count(); i++)
                        {
                            lt_RESUMEN_VP.CurrentIndex = i;
                            resumenvp = new RESUMEN_VP();
                            resumenvp.LAND = lt_RESUMEN_VP.GetString("LAND");
                            resumenvp.ID_CAJA = lt_RESUMEN_VP.GetString("ID_CAJA");
                            resumenvp.SOCIEDAD = lt_RESUMEN_VP.GetString("SOCIEDAD");
                            resumenvp.SOCIEDAD_TXT = lt_RESUMEN_VP.GetString("SOCIEDAD_TXT");
                            resumenvp.VIA_PAGO = lt_RESUMEN_VP.GetString("VIA_PAGO");
                            resumenvp.TEXT1 = lt_RESUMEN_VP.GetString("TEXT1");
                            resumenvp.MONEDA = lt_RESUMEN_VP.GetString("MONEDA");
                            resumenvp.MONTO = lt_RESUMEN_VP.GetString("MONTO");
                            resumenvp.CANT_DOCS = lt_RESUMEN_VP.GetString("CANT_DOCS");
                            resumen_viapago.Add(resumenvp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_DETALLE_REND = BapiGetUser.GetTable("DETALLE_ARQUEO"); //Detalle de efectivo
                    try
                    {
                        for (int i = 0; i < lt_DETALLE_REND.Count(); i++)
                        {
                            lt_DETALLE_REND.CurrentIndex = i;
                            detallerend = new DETALLE_ARQUEO();
                            detallerend.LAND = lt_DETALLE_REND.GetString("LAND");
                            detallerend.ID_CAJA = lt_DETALLE_REND.GetString("ID_CAJA");
                            detallerend.USUARIO = lt_DETALLE_REND.GetString("USUARIO");
                            detallerend.SOCIEDAD = lt_DETALLE_REND.GetString("SOCIEDAD");
                            detallerend.FECHA_REND = lt_DETALLE_REND.GetString("FECHA_REND");
                            detallerend.HORA_REND = lt_DETALLE_REND.GetString("HORA_REND");
                            detallerend.MONEDA = lt_DETALLE_REND.GetString("MONEDA");
                            detallerend.VIA_PAGO = lt_DETALLE_REND.GetString("VIA_PAGO"); //Efectivo
                            detallerend.TIPO_MONEDA = lt_DETALLE_REND.GetString("TIPO_MONEDA"); //Denominacion
                            detallerend.CANTIDAD_MON = lt_DETALLE_REND.GetString("CANTIDAD_MON"); //Cuantos??
                            detallerend.SUMA_MON_BILL = lt_DETALLE_REND.GetString("SUMA_MON_BILL"); //Cantidd*denominacion
                            detallerend.CANTIDAD_DOC = lt_DETALLE_REND.GetString("CANTIDAD_DOC"); //""
                            detallerend.SUMA_DOCS = lt_DETALLE_REND.GetString("SUMA_DOCS");//""
                            detalle_rend.Add(detallerend);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    ls_RETORNO = BapiGetUser.GetTable("RETORNO");
                    try
                    {
                        for (int i = 0; i < ls_RETORNO.Count(); i++)
                        {
                            ls_RETORNO.CurrentIndex = i;
                            retorno = new ESTATUS();
                            if (ls_RETORNO.GetString("TYPE") == "S")
                            {
                                message = message + " - " + ls_RETORNO.GetString("MESSAGE");
                                if (id_arqueo == "")
                                {
                                    id_arqueo = ls_RETORNO.GetString("MESSAGE_V1");
                                }
                            }
                            if (ls_RETORNO.GetString("TYPE") == "E")
                            {
                                errormessage = errormessage + " - " + ls_RETORNO.GetString("MESSAGE");
                            }
                            retorno.TYPE = ls_RETORNO.GetString("TYPE");
                            retorno.ID = ls_RETORNO.GetString("ID");
                            retorno.NUMBER = ls_RETORNO.GetString("NUMBER");
                            retorno.MESSAGE = ls_RETORNO.GetString("MESSAGE");
                            retorno.LOG_NO = ls_RETORNO.GetString("LOG_NO");
                            retorno.LOG_MSG_NO = ls_RETORNO.GetString("LOG_MSG_NO");
                            retorno.MESSAGE_V1 = ls_RETORNO.GetString("MESSAGE_V1");
                            retorno.MESSAGE_V2 = ls_RETORNO.GetString("MESSAGE_V2");
                            retorno.MESSAGE_V3 = ls_RETORNO.GetString("MESSAGE_V3");
                            if (ls_RETORNO.GetString("MESSAGE_V4") != "")
                            {
                            }
                            retorno.MESSAGE_V4 = ls_RETORNO.GetString("MESSAGE_V4");
                            retorno.PARAMETER = ls_RETORNO.GetString("PARAMETER");
                            retorno.ROW = ls_RETORNO.GetString("ROW");
                            retorno.FIELD = ls_RETORNO.GetString("FIELD");
                            retorno.SYSTEM = ls_RETORNO.GetString("SYSTEM");
                            T_Retorno.Add(retorno);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                }
                else
                {
                    errormessage = retval;
                    MessageBox.Show("No se pudo conectar a la RFC");
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, ex.StackTrace);
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }
    }
}
