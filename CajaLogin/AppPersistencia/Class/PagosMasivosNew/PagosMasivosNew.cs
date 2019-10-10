using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;

namespace CajaIndigo.AppPersistencia.Class.PagosMasivosNew
{
    class PagosMasivosNew 
    {

        public List<ESTATUS> objReturn2 = new List<ESTATUS>();
        public string message = "";
        public string errormessage = "";
        public string comprobante = string.Empty;

        ConexSAP connectorSap = new ConexSAP();

        public void pagosmasivos(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA
            , string P_LAND, string P_FECHA, string P_FILE, string P_ID_APERTURA, string P_ID_CAJA, string P_PAY_CURRENCY, List<PagosMasivosNuevo> ListaExc, List<VIAS_PAGO_MASIVO> viasPagoMasivos)
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

                    IRfcTable DetalleViasPago = BapiGetUser.GetTable("VIAS_PAGO_MASIVO");

                    for (var i = 0; i < viasPagoMasivos.Count; i++)
                    {
                        DetalleViasPago.Append();
                        DetalleViasPago.SetValue("MANDT", viasPagoMasivos[i].MANDT);
                        DetalleViasPago.SetValue("LAND", viasPagoMasivos[i].LAND);
                        DetalleViasPago.SetValue("ID_COMPROBANTE", viasPagoMasivos[i].ID_COMPROBANTE);
                        DetalleViasPago.SetValue("ID_DETALLE", viasPagoMasivos[i].ID_DETALLE);
                        DetalleViasPago.SetValue("ID_CAJA", viasPagoMasivos[i].ID_CAJA);
                        DetalleViasPago.SetValue("VIA_PAGO", viasPagoMasivos[i].VIA_PAGO);
                        DetalleViasPago.SetValue("MONTO", viasPagoMasivos[i].MONTO);
                        DetalleViasPago.SetValue("MONEDA", viasPagoMasivos[i].MONEDA);
                        DetalleViasPago.SetValue("BANCO", viasPagoMasivos[i].BANCO);
                        DetalleViasPago.SetValue("EMISOR", viasPagoMasivos[i].EMISOR);
                        DetalleViasPago.SetValue("NUM_CHEQUE", viasPagoMasivos[i].NUM_CHEQUE);
                        DetalleViasPago.SetValue("COD_AUTORIZACION", viasPagoMasivos[i].COD_AUTORIZACION);
                        DetalleViasPago.SetValue("NUM_CUOTAS", viasPagoMasivos[i].NUM_CUOTAS);
                        DetalleViasPago.SetValue("FECHA_VENC", Convert.ToDateTime(viasPagoMasivos[i].FECHA_VENC));
                        DetalleViasPago.SetValue("TEXTO_POSICION", viasPagoMasivos[i].TEXTO_POSICION);
                        DetalleViasPago.SetValue("ANEXO", viasPagoMasivos[i].ANEXO);
                        DetalleViasPago.SetValue("SUCURSAL", viasPagoMasivos[i].SUCURSAL);
                        DetalleViasPago.SetValue("NUM_CUENTA", viasPagoMasivos[i].NUM_CUENTA);
                        DetalleViasPago.SetValue("NUM_TARJETA", viasPagoMasivos[i].NUM_TARJETA);
                        DetalleViasPago.SetValue("NUM_VALE_VISTA", viasPagoMasivos[i].NUM_VALE_VISTA);
                        DetalleViasPago.SetValue("PATENTE", viasPagoMasivos[i].PATENTE);
                        DetalleViasPago.SetValue("NUM_VENTA", viasPagoMasivos[i].NUM_VENTA);
                        DetalleViasPago.SetValue("PAGARE", viasPagoMasivos[i].PAGARE);
                        DetalleViasPago.SetValue("FECHA_EMISION", Convert.ToDateTime(viasPagoMasivos[i].FECHA_EMISION));
                        DetalleViasPago.SetValue("NOMBRE_GIRADOR", viasPagoMasivos[i].NOMBRE_GIRADOR);
                        DetalleViasPago.SetValue("CARTA_CURSE", viasPagoMasivos[i].CARTA_CURSE);
                        DetalleViasPago.SetValue("NUM_TRANSFER", viasPagoMasivos[i].NUM_TRANSFER);
                        DetalleViasPago.SetValue("NUM_DEPOSITO", viasPagoMasivos[i].NUM_DEPOSITO);
                        DetalleViasPago.SetValue("CTA_BANCO", viasPagoMasivos[i].CTA_BANCO);
                        DetalleViasPago.SetValue("IFINAN", viasPagoMasivos[i].IFINAN);
                        DetalleViasPago.SetValue("ZUONR", viasPagoMasivos[i].ZUONR);
                        DetalleViasPago.SetValue("CORRE", viasPagoMasivos[i].CORRE);
                        DetalleViasPago.SetValue("HKONT", viasPagoMasivos[i].HKONT);
                        DetalleViasPago.SetValue("PRCTR", viasPagoMasivos[i].PRCTR);
                        DetalleViasPago.SetValue("ZNOP", viasPagoMasivos[i].ZNOP);
                     }
                    BapiGetUser.SetValue("VIAS_PAGO_MASIVO", DetalleViasPago);


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
                        if (retorno[i].GetString("MESSAGE_V4") != "")
                        {
                            comprobante = retorno[i].GetString("MESSAGE_V4");
                        }
                        //p_return.MESSAGE_V4 = retorno[i].GetString("MESSAGE_V4");
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