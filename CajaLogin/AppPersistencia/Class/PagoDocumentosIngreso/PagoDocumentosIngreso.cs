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
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes.Estructura;
using CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes;
using System.Globalization;
using CajaIndigo.AppPersistencia.Class.PagoDocumentosIngreso.Estructura;
using System.Windows.Forms;

namespace CajaIndigo.AppPersistencia.Class.PagoDocumentosIngreso
{
    class PagoDocumentosIngreso 
    {
       
        public string pagomessage = "";
        public string status = "";
        public string comprobante = "";
        public string message = "";
        public string stringRfc = "";
        public int id_error = 0;
        ConexSAP connectorSap = new ConexSAP();
        public List<ESTATUS> T_Retorno = new List<ESTATUS>();
        public double ValorConvertido;
        public string Monto;

        public List<VALIDAREFECTIVO> validar = new List<VALIDAREFECTIVO>();

        public void pagodocumentosingreso(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_SOCIEDAD, List<DetalleViasPago> P_VIASPAGO, List<T_DOCUMENTOS> P_DOCSAPAGAR, string P_PAIS, string P_MONEDA, string P_CAJA, string P_CAJERO, string P_INGRESO, string P_APAGAR)
        {
            try
            {

                T_Retorno.Clear();
                pagomessage = "";
                status = "";
                comprobante = "";
                message = "";
                stringRfc = "";
                IRfcTable lt_PAGO_DOCS;
                IRfcTable lt_PAGO_MESS;
                // CERR_CAJA CERR_CAJA_resp;
                ESTATUS retorno;
                T_Retorno.Clear();
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_REC_Y_FAC");
                    BapiGetUser.SetValue("ID_CAJA",P_CAJA);
                    BapiGetUser.SetValue("PAY_CURRENCY", P_MONEDA);
                    BapiGetUser.SetValue("LAND", P_PAIS);

                    BapiGetUser.SetValue("TOTAL_FACTURAS", Convert.ToDouble(P_APAGAR));
                    BapiGetUser.SetValue("TOTAL_VIAS",  Convert.ToDouble(P_INGRESO));
                    P_INGRESO = P_INGRESO.Replace(",", "");
                    P_INGRESO = P_INGRESO.Replace(".", "");
                    double Diferencia = Convert.ToDouble(P_APAGAR) - Convert.ToDouble(P_INGRESO);
                    BapiGetUser.SetValue("DIFERENCIA",  Convert.ToDouble(P_APAGAR) - Convert.ToDouble(P_INGRESO));            
                    IRfcTable GralDat3 = BapiGetUser.GetTable("RETURN");
                    try
                    {
                        if (Diferencia != 0)
                        {
                            GralDat3.Append();
                            GralDat3.SetValue("TYPE", "X");
                            GralDat3.SetValue("ID","");
                            GralDat3.SetValue("NUMBER", "");
                            GralDat3.SetValue("MESSAGE", "");
                            GralDat3.SetValue("LOG_NO", "");
                            GralDat3.SetValue("LOG_MSG_NO", "");
                            GralDat3.SetValue("MESSAGE_V1", "");
                            GralDat3.SetValue("MESSAGE_V2", "");
                            GralDat3.SetValue("MESSAGE_V3", "");
                            GralDat3.SetValue("MESSAGE_V4", "");
                            GralDat3.SetValue("PARAMETER", "");
                            GralDat3.SetValue("ROW","");
                            GralDat3.SetValue("FIELD", "");
                            GralDat3.SetValue("SYSTEM","");
                        }
                   }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                    }
                    BapiGetUser.SetValue("RETURN", GralDat3);

                    IRfcTable GralDat = BapiGetUser.GetTable("VIAS_PAGO");
                    try
                    {                    
                        for (var i = 0; i < P_VIASPAGO.Count; i++)
                        {
                            GralDat.Append();
                            GralDat.SetValue("MANDT", P_VIASPAGO[i].MANDT);
                            GralDat.SetValue("LAND", P_VIASPAGO[i].LAND);
                            GralDat.SetValue("ID_COMPROBANTE", P_VIASPAGO[i].ID_COMPROBANTE);
                            GralDat.SetValue("ID_DETALLE", P_VIASPAGO[i].ID_DETALLE);
                            GralDat.SetValue("VIA_PAGO", P_VIASPAGO[i].VIA_PAGO);
                            double Monto = Convert.ToDouble(P_VIASPAGO[i].MONTO); // 100;
                            if (P_VIASPAGO[i].MONEDA == "CLP")
                            {
                                Monto = Monto / 100;
                                GralDat.SetValue("MONTO",Convert.ToString(Monto));
                            }
                            else
                            {
                                GralDat.SetValue("MONTO", P_VIASPAGO[i].MONTO);
                            }
                            GralDat.SetValue("MONEDA", P_VIASPAGO[i].MONEDA);
                            if (P_VIASPAGO[i].BANCO != "")
                            {
                                GralDat.SetValue("BANCO", P_VIASPAGO[i].BANCO.Substring(0, 3));
                            }
                            else
                            {
                                GralDat.SetValue("BANCO", P_VIASPAGO[i].BANCO);
                            }
                            GralDat.SetValue("EMISOR", P_VIASPAGO[i].EMISOR);
                            GralDat.SetValue("NUM_CHEQUE", P_VIASPAGO[i].NUM_CHEQUE);
                            GralDat.SetValue("COD_AUTORIZACION", P_VIASPAGO[i].COD_AUTORIZACION);
                            GralDat.SetValue("NUM_CUOTAS", P_VIASPAGO[i].NUM_CUOTAS);
                            GralDat.SetValue("FECHA_VENC", Convert.ToDateTime(P_VIASPAGO[i].FECHA_VENC));
                            GralDat.SetValue("TEXTO_POSICION", P_VIASPAGO[i].TEXTO_POSICION);
                            GralDat.SetValue("ANEXO", P_VIASPAGO[i].ANEXO);
                            GralDat.SetValue("SUCURSAL", P_VIASPAGO[i].SUCURSAL);
                            GralDat.SetValue("NUM_CUENTA", P_VIASPAGO[i].NUM_CUENTA);
                            GralDat.SetValue("NUM_TARJETA", P_VIASPAGO[i].NUM_TARJETA);
                            GralDat.SetValue("NUM_VALE_VISTA", P_VIASPAGO[i].NUM_VALE_VISTA);
                            GralDat.SetValue("PATENTE", P_VIASPAGO[i].PATENTE);
                            GralDat.SetValue("NUM_VENTA", P_VIASPAGO[i].NUM_VENTA);
                            GralDat.SetValue("PAGARE", P_VIASPAGO[i].PAGARE);
                            GralDat.SetValue("FECHA_EMISION", Convert.ToDateTime(P_VIASPAGO[i].FECHA_EMISION));
                            GralDat.SetValue("NOMBRE_GIRADOR", P_VIASPAGO[i].NOMBRE_GIRADOR);
                            GralDat.SetValue("CARTA_CURSE", P_VIASPAGO[i].CARTA_CURSE);
                            GralDat.SetValue("NUM_TRANSFER", P_VIASPAGO[i].NUM_TRANSFER);
                            GralDat.SetValue("NUM_DEPOSITO", P_VIASPAGO[i].NUM_DEPOSITO);
                            GralDat.SetValue("CTA_BANCO", P_VIASPAGO[i].CTA_BANCO);
                            GralDat.SetValue("IFINAN", P_VIASPAGO[i].IFINAN);
                            GralDat.SetValue("CORRE", P_VIASPAGO[i].CORRE);
                            GralDat.SetValue("ZUONR", P_VIASPAGO[i].ZUONR);
                            GralDat.SetValue("HKONT", P_VIASPAGO[i].HKONT);
                            GralDat.SetValue("PRCTR", P_VIASPAGO[i].PRCTR);
                            GralDat.SetValue("ZNOP", P_VIASPAGO[i].ZNOP);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                    }
                    BapiGetUser.SetValue("VIAS_PAGO", GralDat);

                    IRfcTable GralDat2 = BapiGetUser.GetTable("DOCUMENTOS");
                    try
                    {
                        for (var i = 0; i < P_DOCSAPAGAR.Count; i++)
                        {
                            GralDat2.Append();
                            GralDat2.SetValue("MANDT", "");
                            GralDat2.SetValue("LAND",P_PAIS);
                            GralDat2.SetValue("ID_COMPROBANTE", "");
                            GralDat2.SetValue("POSICION", "");
                            GralDat2.SetValue("CLIENTE", P_DOCSAPAGAR[i].RUTCLI);
                            GralDat2.SetValue("TIPO_DOCUMENTO", P_DOCSAPAGAR[i].CLASE_DOC);
                            GralDat2.SetValue("SOCIEDAD", P_DOCSAPAGAR[i].SOCIEDAD);
                            GralDat2.SetValue("NRO_DOCUMENTO", P_DOCSAPAGAR[i].NDOCTO);
                            GralDat2.SetValue("NRO_REFERENCIA", P_DOCSAPAGAR[i].NREF);
                            GralDat2.SetValue("CAJERO_RESP", P_CAJERO);
                            GralDat2.SetValue("CAJERO_GEN", "");
                            GralDat2.SetValue("ID_CAJA",P_CAJA);
                            GralDat2.SetValue("NRO_COMPENSACION", "");
                            GralDat2.SetValue("TEXTO_CABECERA","");
                            GralDat2.SetValue("NULO", "");
                            GralDat2.SetValue("USR_ANULADOR", "");
                            GralDat2.SetValue("NRO_ANULACION", "");
                            GralDat2.SetValue("APROBADOR_ANULA", "");
                            GralDat2.SetValue("TXT_ANULACION", "");
                            GralDat2.SetValue("EXCEPCION", "");
                            GralDat2.SetValue("FECHA_DOC", Convert.ToDateTime(P_DOCSAPAGAR[i].FECHA_DOC));
                            GralDat2.SetValue("FECHA_VENC_DOC",Convert.ToDateTime(P_DOCSAPAGAR[i].FECVENCI));
                            GralDat2.SetValue("NUM_CUOTA", "");
                            GralDat2.SetValue("MONTO_DOC", P_DOCSAPAGAR[i].MONTO.Trim());
                            GralDat2.SetValue("MONTO_DIFERENCIA", 0);
                            GralDat2.SetValue("TEXTO_EXCEPCION", "");
                            GralDat2.SetValue("PARCIAL", "");
                            GralDat2.SetValue("APROBADOR_EX","");
                            GralDat2.SetValue("MONEDA",P_DOCSAPAGAR[i].MONEDA.Trim());
                            GralDat2.SetValue("CLASE_CUENTA", "D");
                            GralDat2.SetValue("CLASE_DOC", P_DOCSAPAGAR[i].CLASE_DOC);
                            GralDat2.SetValue("NUM_CANCELACION","");
                            GralDat2.SetValue("CME", P_DOCSAPAGAR[i].CME);
                            GralDat2.SetValue("NOTA_VENTA","");
                            GralDat2.SetValue("CEBE", P_DOCSAPAGAR[i].CEBE);
                            GralDat2.SetValue("ACC", P_DOCSAPAGAR[i].ACC);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                    }
                    BapiGetUser.SetValue("DOCUMENTOS", GralDat2);

                    BapiGetUser.Invoke(SapRfcDestination);

                    //LLenamos los datos que retorna la estructura de la RFC
                    //pagomessage = BapiGetUser.GetString("E_MSJ");
                    //id_error = BapiGetUser.GetInt("E_ID_MSJ");
                    //message = BapiGetUser.GetString("E_AUGBL");

                    lt_PAGO_DOCS = BapiGetUser.GetTable("RETURN");
                    for (int i = 0; i < lt_PAGO_DOCS.Count(); i++)
                    {
                        lt_PAGO_DOCS.CurrentIndex = i;
                        retorno = new ESTATUS();
                        if (lt_PAGO_DOCS.GetString("TYPE") == "S")
                        {
                            message = message + " - " + lt_PAGO_DOCS.GetString("MESSAGE") + "\n";
                        }
                         if (lt_PAGO_DOCS.GetString("TYPE") == "E")
                        {
                            pagomessage = pagomessage + " - " + lt_PAGO_DOCS.GetString("MESSAGE") + "\n";
                        }
                        retorno.TYPE = lt_PAGO_DOCS.GetString("TYPE");
                        retorno.ID = lt_PAGO_DOCS.GetString("ID");
                        retorno.NUMBER = lt_PAGO_DOCS.GetString("NUMBER");
                        retorno.MESSAGE = lt_PAGO_DOCS.GetString("MESSAGE");
                        retorno.LOG_NO = lt_PAGO_DOCS.GetString("LOG_NO");
                        retorno.LOG_MSG_NO = lt_PAGO_DOCS.GetString("LOG_MSG_NO");
                        retorno.MESSAGE_V1 = lt_PAGO_DOCS.GetString("MESSAGE_V1");
                        retorno.MESSAGE_V2 = lt_PAGO_DOCS.GetString("MESSAGE_V2");
                        retorno.MESSAGE_V3 = lt_PAGO_DOCS.GetString("MESSAGE_V3");
                        if (lt_PAGO_DOCS.GetString("MESSAGE_V4") != "")
                        {
                            comprobante = lt_PAGO_DOCS.GetString("MESSAGE_V4");
                        }
                        retorno.MESSAGE_V4 = lt_PAGO_DOCS.GetString("MESSAGE_V4");
                        retorno.PARAMETER = lt_PAGO_DOCS.GetString("PARAMETER");
                        retorno.ROW = lt_PAGO_DOCS.GetString("ROW");
                        retorno.FIELD = lt_PAGO_DOCS.GetString("FIELD");
                        retorno.SYSTEM = lt_PAGO_DOCS.GetString("SYSTEM");
                        T_Retorno.Add(retorno);
                    }
                }
                GC.Collect();
            }                                                                                                      
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }


        public void ValidarEfectivo(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_SOCIEDAD, string P_PAIS, string P_CAJA, string P_CAJERO, string ViaPago, string IdApertura)
        {
            IRfcTable lt_RESUMEN_VP;
            VALIDAREFECTIVO valEfec;

            try
            {
                connectorSap.idioma = P_IDIOMA;
                connectorSap.idSistema = P_IDSISTEMA;
                connectorSap.instancia = P_INSTANCIA;
                connectorSap.mandante = P_MANDANTE;
                connectorSap.paswr = P_PASSWORD;
                connectorSap.sapRouter = P_SAPROUTER;
                connectorSap.user = P_UNAME;
                connectorSap.server = P_SERVER;

                string retval = connectorSap.connectionsSAP();

                if (string.IsNullOrEmpty(retval))
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_MON_EFEC");
                    BapiGetUser.SetValue("ID_CAJA", P_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_CAJERO);
                    BapiGetUser.SetValue("ID_APERTURA", IdApertura);

                    BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                    BapiGetUser.SetValue("VIA_PAGO", ViaPago);

                    BapiGetUser.Invoke(SapRfcDestination);

                    lt_RESUMEN_VP = BapiGetUser.GetTable("RESUMEN_VP");

                    for (int i = 0; i < lt_RESUMEN_VP.Count(); i++)
                    {

                        lt_RESUMEN_VP.CurrentIndex = i;
                        valEfec = new VALIDAREFECTIVO();

                        valEfec.LAND = lt_RESUMEN_VP[i].GetString("LAND");
                        valEfec.ID_CAJA = lt_RESUMEN_VP[i].GetString("ID_CAJA");
                        valEfec.SOCIEDAD = lt_RESUMEN_VP[i].GetString("SOCIEDAD");
                        valEfec.SOCIEDAD_TXT = lt_RESUMEN_VP[i].GetString("SOCIEDAD_TXT");
                        valEfec.VIA_PAGO = lt_RESUMEN_VP[i].GetString("VIA_PAGO");
                        valEfec.TEXT1 = lt_RESUMEN_VP[i].GetString("TEXT1");
                        valEfec.MONEDA = lt_RESUMEN_VP[i].GetString("MONEDA");
                        valEfec.MONTO = lt_RESUMEN_VP[i].GetString("MONTO");
                        valEfec.CANT_DOCS = lt_RESUMEN_VP[i].GetString("CANT_DOCS");
                        validar.Add(valEfec);
                    }
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.Write(e.StackTrace);
                throw new Exception();
            }
            finally
            {
                lt_RESUMEN_VP = null;
                valEfec = null;
            }
        }

        public void Conversion(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string LAND, string VPAGO, string FCURR, string TCURR, string FAMOUNT, string TAMOUNT)
        {
            try
            {
                connectorSap.idioma = P_IDIOMA;
                connectorSap.idSistema = P_IDSISTEMA;
                connectorSap.instancia = P_INSTANCIA;
                connectorSap.mandante = P_MANDANTE;
                connectorSap.paswr = P_PASSWORD;
                connectorSap.sapRouter = P_SAPROUTER;
                connectorSap.user = P_UNAME;
                connectorSap.server = P_SERVER;

                string retval = connectorSap.connectionsSAP();

                if (string.IsNullOrEmpty(retval))
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_TCURR");
                    BapiGetUser.SetValue("LAND", LAND);
                    BapiGetUser.SetValue("VPAGO", VPAGO);
                    BapiGetUser.SetValue("FCURR", FCURR);
                    BapiGetUser.SetValue("TCURR", TCURR);
                    BapiGetUser.SetValue("FAMOUNT", FAMOUNT);
                    BapiGetUser.SetValue("TAMOUNT", TAMOUNT);


                    BapiGetUser.Invoke(SapRfcDestination);
               
                    string Val = BapiGetUser.GetValue("LAMOUNT").ToString();
                    double val4 = Math.Ceiling(Convert.ToDouble(Val));
                    ValorConvertido = val4;
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.Write(e.StackTrace);
                throw new Exception();
            }
            finally
            {

            }
        }

        public void Conversion2(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string LAND, string VPAGO, string FCURR, string TCURR, string FAMOUNT, string TAMOUNT)
        {
            try
            {
                connectorSap.idioma = P_IDIOMA;
                connectorSap.idSistema = P_IDSISTEMA;
                connectorSap.instancia = P_INSTANCIA;
                connectorSap.mandante = P_MANDANTE;
                connectorSap.paswr = P_PASSWORD;
                connectorSap.sapRouter = P_SAPROUTER;
                connectorSap.user = P_UNAME;
                connectorSap.server = P_SERVER;

                string retval = connectorSap.connectionsSAP();

                if (string.IsNullOrEmpty(retval))
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_TCURR");
                    BapiGetUser.SetValue("LAND", LAND);
                    BapiGetUser.SetValue("VPAGO", VPAGO);
                    BapiGetUser.SetValue("FCURR", FCURR);
                    BapiGetUser.SetValue("TCURR", TCURR);
                    BapiGetUser.SetValue("FAMOUNT", FAMOUNT);
                    BapiGetUser.SetValue("TAMOUNT", TAMOUNT);


                    BapiGetUser.Invoke(SapRfcDestination);

                    string Val = BapiGetUser.GetValue("LAMOUNT").ToString().Replace(".", ",");
                    double val4 = Math.Ceiling(Convert.ToDouble(Val));
                    ValorConvertido = val4;
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.Write(e.StackTrace);
                throw new Exception();
            }
            finally
            {

            }
        }    
    }
}