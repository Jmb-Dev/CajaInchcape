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
using CajaIndu.AppPersistencia.Class.RecaudacionVehiculos.Estructura;

namespace CajaIndu.AppPersistencia.Class.RecaudacionVehiculos
{
    class Recaudacion_vehi
    {
        public List<IT_PAGOS_CAB> objPagCab = new List<IT_PAGOS_CAB>();
        public List<IT_PAGOS> objPag = new List<IT_PAGOS>();
        public List<RETURN> objReturn2 = new List<RETURN>();
        public string DOCUMENTO;
        public string DOCUMENTO2;
        public string message = "";
        public string errormessage = "";
        ConexSAP connectorSap = new ConexSAP();

        public void recauVehi(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string VBELN, string STCD1, string SOCIEDAD)
        {
            objPag.Clear();
            objPagCab.Clear();
            objReturn2.Clear();
            errormessage = "";
            message = "";
          try
             {
                RETURN p_return;
                IRfcTable lt_IT_PAGOS_CAB;
                IRfcTable lt_IT_PAGOS;

               IT_PAGOS GET_IT_PAGOS_resp;
               IT_PAGOS_CAB GET_IT_PAGOS_CAB_resp;
               FormatoMonedas FM = new FormatoMonedas();

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_RECAU_VEHI");
                    BapiGetUser.SetValue("VBELN", VBELN);
                    BapiGetUser.SetValue("STCD1", STCD1);
                    BapiGetUser.SetValue("SOCIEDAD", SOCIEDAD);

                    BapiGetUser.Invoke(SapRfcDestination);

                    lt_IT_PAGOS_CAB = BapiGetUser.GetTable("IT_PAGOS_CAB");
                    lt_IT_PAGOS = BapiGetUser.GetTable("IT_PAGOS");

                    if (lt_IT_PAGOS_CAB.RowCount > 0)
                    {
                        for (int i = 0; i < lt_IT_PAGOS_CAB.RowCount; i++)
                        {
                            lt_IT_PAGOS_CAB.CurrentIndex = i;
                            GET_IT_PAGOS_CAB_resp = new IT_PAGOS_CAB();

                            GET_IT_PAGOS_CAB_resp.VBELN = lt_IT_PAGOS_CAB[i].GetString("VBELN");
                            GET_IT_PAGOS_CAB_resp.LICPL = lt_IT_PAGOS_CAB[i].GetString("LICPL");
                            GET_IT_PAGOS_CAB_resp.H_NETWR = lt_IT_PAGOS_CAB[i].GetString("H_NETWR");
                            GET_IT_PAGOS_CAB_resp.WAERK = lt_IT_PAGOS_CAB[i].GetString("WAERK");
                            GET_IT_PAGOS_CAB_resp.BUKRS_VF = lt_IT_PAGOS_CAB[i].GetString("BUKRS_VF");
                            GET_IT_PAGOS_CAB_resp.KKBER = lt_IT_PAGOS_CAB[i].GetString("KKBER");
                            GET_IT_PAGOS_CAB_resp.STCD1 = lt_IT_PAGOS_CAB[i].GetString("STCD1");
                            objPagCab.Add(GET_IT_PAGOS_CAB_resp);
                        }
                    }

                    if (lt_IT_PAGOS.RowCount > 0)
                    {
                        for (int i = 0; i < lt_IT_PAGOS.RowCount; i++)
                        {

                            lt_IT_PAGOS.CurrentIndex = i;
                            GET_IT_PAGOS_resp = new IT_PAGOS();

                            GET_IT_PAGOS_resp.VBELN = lt_IT_PAGOS[i].GetString("VBELN");
                            GET_IT_PAGOS_resp.CORRE = lt_IT_PAGOS[i].GetString("CORRE");
                            GET_IT_PAGOS_resp.VIADP = lt_IT_PAGOS[i].GetString("VIADP");
                            GET_IT_PAGOS_resp.DESCV = lt_IT_PAGOS[i].GetString("DESCV");
                            GET_IT_PAGOS_resp.DBM_LICEXT = lt_IT_PAGOS[i].GetString("DBM_LICEXT");
                            GET_IT_PAGOS_resp.NUDOC = lt_IT_PAGOS[i].GetString("NUDOC");
                            GET_IT_PAGOS_resp.CODBA = lt_IT_PAGOS[i].GetString("CODBA");
                            GET_IT_PAGOS_resp.NOMBA = lt_IT_PAGOS[i].GetString("NOMBA");
                            GET_IT_PAGOS_resp.CODIN = lt_IT_PAGOS[i].GetString("CODIN");
                            GET_IT_PAGOS_resp.NOMIN = lt_IT_PAGOS[i].GetString("NOMIN");
                            //string VALORUN = Convert.ToString(lt_IT_PAGOS[i].GetString("MONTO"));
                            //VALORUN = VALORUN.Replace(".", "");
                            //VALORUN = VALORUN.Replace(",", "");
                            //decimal monto = Convert.ToDecimal(VALORUN);
                            //GET_IT_PAGOS_resp.MONTO = string.Format("{0:0,0}", monto);
                            GET_IT_PAGOS_resp.KUNNR = lt_IT_PAGOS[i].GetString("KUNNR");
                            if (lt_IT_PAGOS[i].GetString("WAERS") == "CLP")
                            {
                                GET_IT_PAGOS_resp.MONTO = FM.FormatoMonedaChilena(lt_IT_PAGOS[i].GetString("MONTO"), "2");
                                //paramt.SetValue("MONTO_DOC", FM.FormatoMonedaChilena(DocumPago[i].MONTO_DOC, "2"));
                            }
                            else
                            {
                                GET_IT_PAGOS_resp.MONTO = FM.FormatoMonedaExtranjera(lt_IT_PAGOS[i].GetString("MONTO"));
                                //paramt.SetValue("MONTO_DOC", FM.FormatoMonedaExtranjera(DocumPago[i].MONTO_DOC));
                            }
                            //GET_IT_PAGOS_resp.MONTO = lt_IT_PAGOS[i].GetString("MONTO");
                            GET_IT_PAGOS_resp.CTACE = lt_IT_PAGOS[i].GetString("CTACE");
                            GET_IT_PAGOS_resp.FEACT = lt_IT_PAGOS[i].GetString("FEACT");
                            GET_IT_PAGOS_resp.FEVEN = lt_IT_PAGOS[i].GetString("FEVEN");
                            GET_IT_PAGOS_resp.INTER = lt_IT_PAGOS[i].GetString("INTER");
                            GET_IT_PAGOS_resp.TASAI = lt_IT_PAGOS[i].GetString("TASAI");
                            GET_IT_PAGOS_resp.CUOTA = lt_IT_PAGOS[i].GetString("CUOTA");
                            GET_IT_PAGOS_resp.MINTE = lt_IT_PAGOS[i].GetString("MINTE");
                            //string VALORTOTIN = Convert.ToString(lt_IT_PAGOS[i].GetString("TOTIN"));
                            //VALORTOTIN = VALORTOTIN.Replace(".", "");
                            //VALORTOTIN = VALORTOTIN.Replace(",", "");
                            //decimal TOTIN2 = Convert.ToDecimal(VALORTOTIN);
                            if (lt_IT_PAGOS[i].GetString("WAERS") == "CLP")
                            {
                                GET_IT_PAGOS_resp.TOTIN = FM.FormatoMonedaChilena(lt_IT_PAGOS[i].GetString("TOTIN"), "2");
                                //paramt.SetValue("MONTO_DOC", FM.FormatoMonedaChilena(DocumPago[i].MONTO_DOC, "2"));
                            }
                            else
                            {
                                GET_IT_PAGOS_resp.TOTIN = FM.FormatoMonedaExtranjera(lt_IT_PAGOS[i].GetString("TOTIN"));
                                //paramt.SetValue("MONTO_DOC", FM.FormatoMonedaExtranjera(DocumPago[i].MONTO_DOC));
                            }
                            //GET_IT_PAGOS_resp.TOTIN = string.Format("{0:0,0}", TOTIN2);
                            // GET_IT_PAGOS_resp.TOTIN = lt_IT_PAGOS[i].GetString("TOTIN");
                            GET_IT_PAGOS_resp.RUTGI = lt_IT_PAGOS[i].GetString("RUTGI");
                            GET_IT_PAGOS_resp.NOMGI = lt_IT_PAGOS[i].GetString("NOMGI");
                            GET_IT_PAGOS_resp.WAERS = lt_IT_PAGOS[i].GetString("WAERS");
                            GET_IT_PAGOS_resp.STAT = lt_IT_PAGOS.GetString("STAT");
                            GET_IT_PAGOS_resp.PRCTR = lt_IT_PAGOS[i].GetString("PRCTR");
                            GET_IT_PAGOS_resp.KUNNR = lt_IT_PAGOS[i].GetString("KUNNR");
                            GET_IT_PAGOS_resp.KKBER = lt_IT_PAGOS[i].GetString("KKBER");
                            GET_IT_PAGOS_resp.STCD1 = lt_IT_PAGOS[i].GetString("STCD1");
                            GET_IT_PAGOS_resp.HKONT = lt_IT_PAGOS[i].GetString("HKONT");
                            GET_IT_PAGOS_resp.BANKN = lt_IT_PAGOS[i].GetString("BANKN");
                            objPag.Add(GET_IT_PAGOS_resp);
                        }
                    }
                    else
                    {
                       // MessageBox.Show("No existen datos para este número de documento o RUT");
                    }


                    IRfcTable retorno = BapiGetUser.GetTable("RETURN");

                    for (var i = 0; i < retorno.RowCount; i++)
                    {
                        retorno.CurrentIndex = i;

                        p_return = new RETURN();

                        p_return.TYPE = retorno[i].GetString("TYPE");
                        p_return.ID = retorno[i].GetString("ID");
                        p_return.NUMBER = retorno[i].GetString("NUMBER");
                        p_return.MESSAGE = retorno[i].GetString("MESSAGE");
                        p_return.LOG_NO = retorno[i].GetString("LOG_NO");
                        p_return.LOG_MSG_NO = retorno[i].GetString("LOG_MSG_NO");
                        p_return.MESSAGE_V1 = retorno[i].GetString("MESSAGE_V1");
                        p_return.MESSAGE_V2 = retorno[i].GetString("MESSAGE_V2");
                        p_return.MESSAGE_V3 = retorno[i].GetString("MESSAGE_V3");
                        p_return.MESSAGE_V4 = retorno[i].GetString("MESSAGE_V4");
                        p_return.PARAMETER = retorno[i].GetString("PARAMETER");
                        p_return.ROW = retorno[i].GetString("ROW");
                        p_return.FIELD = retorno[i].GetString("FIELD");
                        p_return.SYSTEM = retorno[i].GetString("SYSTEM");
                        objReturn2.Add(p_return);

                    }

                  
                }
             }
          catch (Exception ex)
          {
              Console.WriteLine("{0} Exception caught.", ex);
          }
        }
        public void PagaVehicu(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string ID_CAJA, string TOTAL_VENTA, List<VIAS_PAGO_VEHI> ViaPago, List<DOCUMENTO_CAB> DocumPago, List<ACT_FPAGOS> pago, string PAY_CURRENCY, string RUT, string SOCIEDAD, string NOTA_VENTA, string TOTAL_VIAS, string LAND) 
        {
            
            objReturn2.Clear();
            DOCUMENTO = "";
            DOCUMENTO2 = "";
            IRfcTable lt_VIAS_PAGO_VEHI;

            VIAS_PAGO_VEHI VIAS_PAGO_VEHI_RESP;
            errormessage = "";
            message = "";
              try
            {
                RETURN p_return;
                TOTAL_VENTA = TOTAL_VENTA.Replace(".", "");
                TOTAL_VENTA = TOTAL_VENTA.Replace(",", "");
                decimal totalventa_d = Convert.ToDecimal(TOTAL_VENTA);
                FormatoMonedas FM = new FormatoMonedas();

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

                if (string.IsNullOrEmpty(retval))
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_PAGO_ANT_VEHI");

                    BapiGetUser.SetValue("PAY_CURRENCY", PAY_CURRENCY);
                    BapiGetUser.SetValue("RUT", RUT);
                    BapiGetUser.SetValue("SOCIEDAD", SOCIEDAD);
                    BapiGetUser.SetValue("NOTA_VENTA", NOTA_VENTA);
                    BapiGetUser.SetValue("TOTAL_VIAS", TOTAL_VIAS);
                    BapiGetUser.SetValue("LAND", LAND);
                    BapiGetUser.SetValue("ID_CAJA", ID_CAJA);
                    BapiGetUser.SetValue("TOTAL_VENTA", totalventa_d);

                   // LLENAMOS TABLA DOCUMENTO
                    IRfcTable paramt = BapiGetUser.GetTable("DOCUMENTO_CAB");
                    for (int i = 0; i < DocumPago.Count(); i++)
                    {
                        paramt.Append();
                        paramt.SetValue("MANDT", DocumPago[i].MANDT);
                        paramt.SetValue("LAND", DocumPago[i].LAND);
                        paramt.SetValue("ID_COMPROBANTE", DocumPago[i].ID_COMPROBANTE);
                        paramt.SetValue("POSICION", DocumPago[i].POSICION);
                        paramt.SetValue("ID_CAJA", DocumPago[i].ID_CAJA);
                        paramt.SetValue("ID_APERTURA", DocumPago[i].ID_APERTURA);
                        paramt.SetValue("CLIENTE", DocumPago[i].CLIENTE);
                        paramt.SetValue("TIPO_DOCUMENTO", DocumPago[i].TIPO_DOCUMENTO);
                        paramt.SetValue("SOCIEDAD", DocumPago[i].SOCIEDAD);
                        paramt.SetValue("NRO_DOCUMENTO", DocumPago[i].NRO_DOCUMENTO);
                        paramt.SetValue("NRO_REFERENCIA", DocumPago[i].NRO_REFERENCIA);
                        paramt.SetValue("CAJERO_RESP", DocumPago[i].CAJERO_RESP);
                        paramt.SetValue("CAJERO_GEN", DocumPago[i].CAJERO_GEN);
                        paramt.SetValue("FECHA_COMP", DocumPago[i].FECHA_COMP);
                        paramt.SetValue("HORA", DocumPago[i].HORA);
                        paramt.SetValue("NRO_COMPENSACION", DocumPago[i].NRO_COMPENSACION);
                        paramt.SetValue("TEXTO_CABECERA", DocumPago[i].TEXTO_CABECERA);
                        paramt.SetValue("NULO", DocumPago[i].NULO);
                        paramt.SetValue("USR_ANULADOR", DocumPago[i].USR_ANULADOR);
                        paramt.SetValue("NRO_ANULACION", DocumPago[i].NRO_ANULACION);
                        paramt.SetValue("APROBADOR_ANULA", DocumPago[i].APROBADOR_ANULA);
                        paramt.SetValue("TXT_ANULACION", DocumPago[i].TXT_ANULACION);
                        paramt.SetValue("EXCEPCION", DocumPago[i].EXCEPCION);
                        paramt.SetValue("FECHA_DOC", DocumPago[i].FECHA_DOC);
                        paramt.SetValue("FECHA_VENC_DOC", DocumPago[i].FECHA_VENC_DOC);
                        paramt.SetValue("NUM_CUOTA", DocumPago[i].NUM_CUOTA);
                        //if (DocumPago[i].MONEDA == "CLP")
                        //{
                        //    paramt.SetValue("MONTO_DOC",FM.FormatoMonedaChilena(DocumPago[i].MONTO_DOC, "2"));
                        //}
                        //else
                        //{
                        //    paramt.SetValue("MONTO_DOC", FM.FormatoMonedaExtranjera(DocumPago[i].MONTO_DOC));
                        //}
                       paramt.SetValue("MONTO_DOC", DocumPago[i].MONTO_DOC);
                        //if (DocumPago[i].MONEDA == "CLP")
                        //{
                        //    paramt.SetValue("MONTO_DIFERENCIA", FM.FormatoMonedaChilena(DocumPago[i].MONTO_DIFERENCIA, "2"));
                        //}
                        //else
                        //{
                        //    paramt.SetValue("MONTO_DIFERENCIA", FM.FormatoMonedaExtranjera(DocumPago[i].MONTO_DIFERENCIA));
                        //}
                        paramt.SetValue("MONTO_DIFERENCIA", DocumPago[i].MONTO_DIFERENCIA);
                        paramt.SetValue("TEXTO_EXCEPCION", DocumPago[i].TEXTO_EXCEPCION);
                        paramt.SetValue("PARCIAL", DocumPago[i].PARCIAL);
                        paramt.SetValue("TIME", DocumPago[i].TIME);
                        paramt.SetValue("APROBADOR_EX", DocumPago[i].APROBADOR_EX);
                        paramt.SetValue("MONEDA", DocumPago[i].MONEDA);
                        paramt.SetValue("CLASE_CUENTA", DocumPago[i].CLASE_CUENTA);
                        paramt.SetValue("CLASE_DOC", DocumPago[i].CLASE_DOC);
                        paramt.SetValue("NUM_CANCELACION", DocumPago[i].NUM_CANCELACION);
                        paramt.SetValue("CME", DocumPago[i].CME);
                        paramt.SetValue("NOTA_VENTA", DocumPago[i].NOTA_VENTA);
                        paramt.SetValue("CEBE", DocumPago[i].CEBE);
                        paramt.SetValue("ACC", DocumPago[i].ACC);

                    }
                  
                    //LLENAMOS TABLA VIAS DE PAGO
                    IRfcTable paramt2 = BapiGetUser.GetTable("VIAS_PAGO_VEHI");

                    for (int i = 0; i < ViaPago.Count(); i++)
                    {
                        paramt2.Append();
                        paramt2.SetValue("MANDT", ViaPago[i].MANDT);
                        paramt2.SetValue("LAND", ViaPago[i].LAND);
                        paramt2.SetValue("ID_COMPROBANTE", ViaPago[i].ID_COMPROBANTE);
                        paramt2.SetValue("ID_DETALLE", ViaPago[i].ID_DETALLE);
                        paramt2.SetValue("ID_CAJA", ViaPago[i].ID_CAJA);
                        paramt2.SetValue("VIA_PAGO", ViaPago[i].VIA_PAGO);
                        //if (ViaPago[i].MONEDA == "CLP")
                        //{
                        //    paramt2.SetValue("MONTO", FM.FormatoMonedaChilena(ViaPago[i].MONTO, "1"));
                        //}
                        //else
                        //{
                        //    paramt2.SetValue("MONTO", FM.FormatoMonedaExtranjera(ViaPago[i].MONTO));
                        //}
                        paramt2.SetValue("MONTO", ViaPago[i].MONTO);
                        paramt2.SetValue("MONEDA", ViaPago[i].MONEDA);
                        paramt2.SetValue("BANCO", ViaPago[i].BANCO);
                        paramt2.SetValue("EMISOR", ViaPago[i].EMISOR);
                        paramt2.SetValue("NUM_CHEQUE", ViaPago[i].NUM_CHEQUE);
                        paramt2.SetValue("COD_AUTORIZACION", ViaPago[i].COD_AUTORIZACION);
                        paramt2.SetValue("NUM_CUOTAS", ViaPago[i].NUM_CUOTAS);
                        paramt2.SetValue("FECHA_VENC", ViaPago[i].FECHA_VENC);
                        paramt2.SetValue("TEXTO_POSICION", ViaPago[i].TEXTO_POSICION);
                        paramt2.SetValue("ANEXO", ViaPago[i].ANEXO);
                        paramt2.SetValue("SUCURSAL", ViaPago[i].SUCURSAL);
                        paramt2.SetValue("NUM_CUENTA", ViaPago[i].NUM_CUENTA);
                        paramt2.SetValue("NUM_TARJETA", ViaPago[i].NUM_TARJETA);
                        paramt2.SetValue("NUM_VALE_VISTA", ViaPago[i].NUM_VALE_VISTA);
                        paramt2.SetValue("PATENTE", ViaPago[i].PATENTE);
                        paramt2.SetValue("NUM_VENTA", ViaPago[i].NUM_VENTA);
                        paramt2.SetValue("PAGARE", ViaPago[i].PAGARE);
                        if ((ViaPago[i].VIA_PAGO == "B") | (ViaPago[i].VIA_PAGO == "U"))
                        {
                            paramt2.SetValue("FECHA_EMISION", Convert.ToDateTime(ViaPago[i].FECHA_EMISION));
                        }
                        else
                        {
                            paramt2.SetValue("FECHA_EMISION", ViaPago[i].FECHA_EMISION);
                        }
                        paramt2.SetValue("NOMBRE_GIRADOR", ViaPago[i].NOMBRE_GIRADOR);
                        paramt2.SetValue("CARTA_CURSE", ViaPago[i].CARTA_CURSE);
                        paramt2.SetValue("NUM_TRANSFER", ViaPago[i].NUM_TRANSFER);
                        paramt2.SetValue("NUM_DEPOSITO", ViaPago[i].NUM_DEPOSITO);
                        paramt2.SetValue("CTA_BANCO", ViaPago[i].CTA_BANCO);
                        paramt2.SetValue("IFINAN", ViaPago[i].IFINAN);
                        paramt2.SetValue("ZUONR", ViaPago[i].ZUONR);
                        paramt2.SetValue("CORRE", ViaPago[i].CORRE);
                        paramt2.SetValue("HKONT", ViaPago[i].HKONT);
                        paramt2.SetValue("PRCTR", ViaPago[i].PRCTR);
                        paramt2.SetValue("ZNOP", ViaPago[i].ZNOP);
                    }

                    IRfcTable paramt3 = BapiGetUser.GetTable("ACT_FPAGOS");
                    for (int i = 0; i < pago.Count(); i++)
                    {
                        paramt3.Append();
                        paramt3.SetValue("VBELN", pago[i].VBELN);
                        paramt3.SetValue("CORRE", pago[i].CORRE);
                    }

                    BapiGetUser.SetValue("DOCUMENTO_CAB", paramt);
                    BapiGetUser.SetValue("VIAS_PAGO_VEHI", paramt2);
                    BapiGetUser.SetValue("ACT_FPAGOS", paramt3);

                    BapiGetUser.Invoke(SapRfcDestination);


                    DOCUMENTO = BapiGetUser.GetValue("DOCUMENTO").ToString();
                    DOCUMENTO2 = BapiGetUser.GetValue("COMPROBANTE").ToString();
                    IRfcTable retorno = BapiGetUser.GetTable("RETORNO");

                    for (var i = 0; i < retorno.RowCount; i++)
                    {
                        retorno.CurrentIndex = i;

                        p_return = new RETURN();

                        p_return.TYPE = retorno[i].GetString("TYPE");
                        p_return.ID = retorno[i].GetString("ID");
                        p_return.NUMBER = retorno[i].GetString("NUMBER");
                        p_return.MESSAGE = retorno[i].GetString("MESSAGE");
                        p_return.LOG_NO = retorno[i].GetString("LOG_NO");
                        p_return.LOG_MSG_NO = retorno[i].GetString("LOG_MSG_NO");
                        p_return.MESSAGE_V1 = retorno[i].GetString("MESSAGE_V1");
                        p_return.MESSAGE_V2 = retorno[i].GetString("MESSAGE_V2");
                        p_return.MESSAGE_V3 = retorno[i].GetString("MESSAGE_V3");
                        p_return.MESSAGE_V4 = retorno[i].GetString("MESSAGE_V4");
                        p_return.PARAMETER = retorno[i].GetString("PARAMETER");
                        p_return.ROW = retorno[i].GetString("ROW");
                        p_return.FIELD = retorno[i].GetString("FIELD");
                        p_return.SYSTEM = retorno[i].GetString("SYSTEM");

                        if (retorno[i].GetString("TYPE") == "E")
                        {
                            if (errormessage.Contains(retorno[i].GetString("MESSAGE")))
                            {
                                ;
                            }
                            else
                            {
                                errormessage = errormessage + " - " + retorno[i].GetString("MESSAGE");
                            }
                        }
                        if (retorno[i].GetString("TYPE") == "S")
                        {
                            if (message.Contains(retorno[i].GetString("MESSAGE")))
                            {
                                ;
                            }
                            else
                            {
                                message = message + " - " + retorno[i].GetString("MESSAGE");
                            }
                        }
                       

                        objReturn2.Add(p_return);

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
            }
        
        }
    }
}