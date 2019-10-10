using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.BusquedaAnulacion.Estructura;
using CajaIndigo.AppPersistencia.Class.Connections;
using SAP.Middleware.Connector;

namespace CajaIndigo.AppPersistencia.Class.BusquedaAnulacion
{
    class BusquedaAnulacion
    {
        public List<CAB_COMP> CabeceraDocs = new List<CAB_COMP>();
        public List<DET_COMP> DetalleDocs = new List<DET_COMP>();
        public List<RETORNO> Retorno = new List<RETORNO>();
        public string errormessage = "";
        ConexSAP connectorSap = new ConexSAP();

        public void docsanulacion(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_DOCUMENTO, string P_RUT,
            string P_SOCIEDAD, string P_LAND, string P_IDCAJA, string P_TP_DOC)
        {
            CabeceraDocs.Clear();
            DetalleDocs.Clear();
            Retorno.Clear();
            IRfcTable lt_h_documentos;
            IRfcTable lt_d_documentos;
            IRfcTable lt_retorno;
            FormatoMonedas FM = new FormatoMonedas();
            //  PART_ABIERTAS  PART_ABIERTAS_resp;
            CAB_COMP DOCS_CABECERA_resp;
            DET_COMP DOCS_DETALLES_resp;
            RETORNO retorno_resp;

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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_BUSCA_COMP_ANULAR");
              
                BapiGetUser.SetValue("ID_COMPROBANTE", P_DOCUMENTO);
                BapiGetUser.SetValue("RUT", P_RUT);
                BapiGetUser.SetValue("LAND", P_LAND);
                BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                BapiGetUser.SetValue("ID_CAJA", P_IDCAJA);
                BapiGetUser.SetValue("TP_DOC", P_TP_DOC);
                BapiGetUser.Invoke(SapRfcDestination);

                lt_h_documentos = BapiGetUser.GetTable("CAB_COMP");
                lt_d_documentos = BapiGetUser.GetTable("DET_COMP");
                lt_retorno = BapiGetUser.GetTable("RETORNO");
                
                if (lt_h_documentos.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_h_documentos.RowCount; i++)
                        {
                            try
                            {
                                lt_h_documentos.CurrentIndex = i;
                                DOCS_CABECERA_resp = new CAB_COMP();
                                DOCS_CABECERA_resp.LAND = P_LAND;
                                DOCS_CABECERA_resp.ID_CAJA = P_IDCAJA;
                                DOCS_CABECERA_resp.ID_COMPROBANTE = lt_h_documentos[i].GetString("ID_COMPROBANTE");
                                DOCS_CABECERA_resp.TIPO_DOCUMENTO = lt_h_documentos[i].GetString("TIPO_DOCUMENTO");
                                DOCS_CABECERA_resp.DESCRIPCION = lt_h_documentos[i].GetString("DESCRIPCION");
                                DOCS_CABECERA_resp.NRO_REFERENCIA = lt_h_documentos[i].GetString("NRO_REFERENCIA");
                                DOCS_CABECERA_resp.FECHA_COMP = lt_h_documentos[i].GetString("FECHA_COMP");
                                DOCS_CABECERA_resp.FECHA_VENC_DOC = lt_h_documentos[i].GetString("FECHA_VENC_DOC");
                                if (lt_h_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                   DOCS_CABECERA_resp.MONTO_DOC = FM.FormatoMonedaChilena(lt_h_documentos[i].GetString("MONTO_DOC").Trim(), "2");
                                }
                                else
                                {
                                    //string moneda = Convert.ToString(lt_h_documentos[i].GetString("MONTO_DOC"));
                                    //decimal ValorAux = Convert.ToDecimal(moneda);
                                    DOCS_CABECERA_resp.MONTO_DOC = FM.FormatoMonedaExtranjera(lt_h_documentos[i].GetString("MONTO_DOC").Trim());
                                }
                                //DOCS_CABECERA_resp.MONTO_DOC = lt_h_documentos[i].GetString("MONTO_DOC");
                                DOCS_CABECERA_resp.TEXTO_EXCEPCION = lt_h_documentos[i].GetString("TEXTO_EXCEPCION");
                                DOCS_CABECERA_resp.CLIENTE = lt_h_documentos[i].GetString("CLIENTE");
                                DOCS_CABECERA_resp.MONEDA = lt_h_documentos[i].GetString("MONEDA");
                                DOCS_CABECERA_resp.CLASE_DOC = lt_h_documentos[i].GetString("CLASE_DOC");
                                DOCS_CABECERA_resp.TXT_CLASE_DOC = lt_h_documentos[i].GetString("TXT_CLASE_DOC");
                                DOCS_CABECERA_resp.NUM_CANCELACION = lt_h_documentos[i].GetString("NUM_CANCELACION");
                                DOCS_CABECERA_resp.AUT_JEF = lt_h_documentos[i].GetString("AUT_JEF");
                                DOCS_CABECERA_resp.VIA_PAGO = lt_h_documentos[i].GetString("VIA_PAGO");
                            CabeceraDocs.Add(DOCS_CABECERA_resp);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                            }

                        }
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No existe(n) registro(s)");
                    }

                    if (lt_d_documentos.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_d_documentos.RowCount; i++)
                        {
                            try
                            {
                                lt_d_documentos.CurrentIndex = i;
                                DOCS_DETALLES_resp = new DET_COMP();

                                DOCS_DETALLES_resp.ID_COMPROBANTE = lt_d_documentos[i].GetString("ID_COMPROBANTE");
                                DOCS_DETALLES_resp.ID_DETALLE = lt_d_documentos[i].GetString("ID_DETALLE");
                                DOCS_DETALLES_resp.VIA_PAGO = lt_d_documentos[i].GetString("VIA_PAGO");
                                DOCS_DETALLES_resp.DESCRIP_VP = lt_d_documentos[i].GetString("DESCRIP_VP");
                                DOCS_DETALLES_resp.NUM_CHEQUE = lt_d_documentos[i].GetString("NUM_CHEQUE");
                                DOCS_DETALLES_resp.FECHA_VENC = lt_d_documentos[i].GetString("FECHA_VENC");
                                if (lt_d_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    DOCS_DETALLES_resp.MONTO = FM.FormatoMonedaChilena(lt_d_documentos[i].GetString("MONTO").Trim(), "1");
                                }
                                else
                                {
                                    DOCS_DETALLES_resp.MONTO = FM.FormatoMonedaExtranjera(lt_d_documentos[i].GetString("MONTO").Trim());
                                }
                                //DOCS_DETALLES_resp.MONTO = lt_d_documentos[i].GetString("MONTO");
                                DOCS_DETALLES_resp.MONEDA = lt_d_documentos[i].GetString("MONEDA");
                                DOCS_DETALLES_resp.NUM_CUOTAS = lt_d_documentos[i].GetString("NUM_CUOTAS");
                                DOCS_DETALLES_resp.EMISOR = lt_d_documentos[i].GetString("EMISOR");
                                DetalleDocs.Add(DOCS_DETALLES_resp);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                            }

                        }
                    }

                        String Mensaje = "";
                    if (lt_retorno.Count > 0)
                    {
                        for (int i = 0; i < lt_retorno.Count(); i++)
                        {
                            lt_retorno.CurrentIndex = i;
                            retorno_resp = new RETORNO();
                            retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                            retorno_resp.ID = lt_retorno.GetString("ID");
                            retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                            retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                            retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                            retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                            retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                            retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                            Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V1");
                            retorno_resp.MESSAGE_V2 = lt_retorno.GetString("MESSAGE_V2");
                            retorno_resp.MESSAGE_V3 = lt_retorno.GetString("MESSAGE_V3");
                            retorno_resp.MESSAGE_V4 = lt_retorno.GetString("MESSAGE_V4");
                            retorno_resp.PARAMETER = lt_retorno.GetString("PARAMETER");
                            retorno_resp.ROW = lt_retorno.GetString("ROW");
                            retorno_resp.FIELD = lt_retorno.GetString("FIELD");
                            retorno_resp.SYSTEM = lt_retorno.GetString("SYSTEM");
                            Retorno.Add(retorno_resp);
                        }
                        System.Windows.MessageBox.Show(Mensaje);
                    }
                    //else
                    //{
                    //    System.Windows.MessageBox.Show("No existe(n) registro(s)");
                    //}

                   
               

            }
            else
            {
                errormessage = retval;
            }
            GC.Collect();
        }
    }
}
    
