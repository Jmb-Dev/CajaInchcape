using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CajaIndigo.AppPersistencia.Class.BusquedaAnulacion.Estructura;
using CajaIndigo.AppPersistencia.Class.Connections;
using SAP.Middleware.Connector;


namespace CajaIndigo.AppPersistencia.Class.CheckUserAnulacion
{
    class CheckUserAnulacion
    {
        public List<CAB_COMP> CabeceraDocs = new List<CAB_COMP>();
        public List<RETORNO> Retorno = new List<RETORNO>();
        public string errormessage = "";
        public string message = "";
        public string valido = "";
        public string estado = "";
        ConexSAP connectorSap = new ConexSAP();

        public void checkdocsanulacion(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE
            , string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_USUARIO, string P_ID_CAJA, string P_LAND, string P_RUT 
            , string P_ID_COMPROBANTE, string P_SOCIEDAD, string P_TP_DOC, List<CAB_COMP> P_CAB_COM)
        {
            CabeceraDocs.Clear();
            Retorno.Clear();
            IRfcTable lt_h_documentos;
          //IRfcTable lt_d_documentos;
            IRfcStructure lt_retorno;

          //PART_ABIERTAS  PART_ABIERTAS_resp;
            CAB_COMP DOCS_CABECERA_resp;
          //DET_COMP DOCS_DETALLES_resp;
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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_CHECK_JEFE_CAJA");
                BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                BapiGetUser.SetValue("USUARIO", P_USUARIO);
                IRfcTable GralDat = BapiGetUser.GetTable("CAB_COMP");

                try
                {
                    for (var i = 0; i < P_CAB_COM.Count; i++)
                    {
                        GralDat.Append();
                        GralDat.SetValue("LAND", P_CAB_COM[i].LAND);
                        GralDat.SetValue("ID_CAJA", P_CAB_COM[i].ID_CAJA);
                        GralDat.SetValue("ID_COMPROBANTE", P_CAB_COM[i].ID_COMPROBANTE);
                        GralDat.SetValue("TIPO_DOCUMENTO", P_CAB_COM[i].TIPO_DOCUMENTO);
                        GralDat.SetValue("DESCRIPCION", P_CAB_COM[i].DESCRIPCION);
                        GralDat.SetValue("NRO_REFERENCIA", P_CAB_COM[i].NRO_REFERENCIA);
                        GralDat.SetValue("FECHA_COMP", P_CAB_COM[i].FECHA_COMP);
                        GralDat.SetValue("FECHA_VENC_DOC", P_CAB_COM[i].FECHA_VENC_DOC);
                        GralDat.SetValue("MONTO_DOC", P_CAB_COM[i].MONTO_DOC);
                        GralDat.SetValue("TEXTO_EXCEPCION", P_CAB_COM[i].TEXTO_EXCEPCION);
                        GralDat.SetValue("CLIENTE", Convert.ToDateTime(P_CAB_COM[i].CLIENTE));
                        GralDat.SetValue("MONEDA", P_CAB_COM[i].MONEDA);
                        GralDat.SetValue("CLASE_DOC", P_CAB_COM[i].CLASE_DOC);
                        GralDat.SetValue("TXT_CLASE_DOC", P_CAB_COM[i].TXT_CLASE_DOC);
                        GralDat.SetValue("NUM_CANCELACION", P_CAB_COM[i].NUM_CANCELACION);
                        GralDat.SetValue("AUT_JEF", P_CAB_COM[i].AUT_JEF);                      
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                }
                BapiGetUser.SetValue("CAB_COMP", GralDat);
          
                BapiGetUser.Invoke(SapRfcDestination);

                valido = BapiGetUser.GetString("VALIDO");
                lt_h_documentos = BapiGetUser.GetTable("CAB_COMP");
                lt_retorno = BapiGetUser.GetStructure("ESTADO");
                
                if (lt_h_documentos.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_h_documentos.RowCount; i++)
                        {
                            try
                            {
                                lt_h_documentos.CurrentIndex = i;
                                DOCS_CABECERA_resp = new CAB_COMP();
                                DOCS_CABECERA_resp.LAND = lt_h_documentos[i].GetString("LAND"); 
                                DOCS_CABECERA_resp.ID_CAJA = lt_h_documentos[i].GetString("ID_CAJA"); 
                                DOCS_CABECERA_resp.ID_COMPROBANTE = lt_h_documentos[i].GetString("ID_COMPROBANTE");
                                DOCS_CABECERA_resp.TIPO_DOCUMENTO = lt_h_documentos[i].GetString("TIPO_DOCUMENTO");
                                DOCS_CABECERA_resp.DESCRIPCION = lt_h_documentos[i].GetString("DESCRIPCION");
                                DOCS_CABECERA_resp.NRO_REFERENCIA = lt_h_documentos[i].GetString("NRO_REFERENCIA");
                                DOCS_CABECERA_resp.FECHA_COMP = lt_h_documentos[i].GetString("FECHA_COMP");
                                DOCS_CABECERA_resp.FECHA_VENC_DOC = lt_h_documentos[i].GetString("FECHA_VENC_DOC");
                                DOCS_CABECERA_resp.MONTO_DOC = lt_h_documentos[i].GetString("MONTO_DOC");
                                DOCS_CABECERA_resp.TEXTO_EXCEPCION = lt_h_documentos[i].GetString("TEXTO_EXCEPCION");
                                DOCS_CABECERA_resp.CLIENTE = lt_h_documentos[i].GetString("CLIENTE");
                                DOCS_CABECERA_resp.MONEDA = lt_h_documentos[i].GetString("MONEDA");
                                DOCS_CABECERA_resp.CLASE_DOC = lt_h_documentos[i].GetString("CLASE_DOC");
                                DOCS_CABECERA_resp.TXT_CLASE_DOC = lt_h_documentos[i].GetString("TXT_CLASE_DOC");
                                DOCS_CABECERA_resp.NUM_CANCELACION = lt_h_documentos[i].GetString("NUM_CANCELACION");
                                DOCS_CABECERA_resp.AUT_JEF = lt_h_documentos[i].GetString("AUT_JEF");
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
                       String Mensaje = "";
                    if (lt_retorno.Count > 0)
                    {
                        for (int i = 0; i < lt_retorno.Count(); i++)
                        {
                            //lt_retorno.CurrentIndex = i;
                            retorno_resp = new RETORNO();
                            retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                            retorno_resp.ID = lt_retorno.GetString("ID");
                            retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                            retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                            retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                            retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                            retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                            retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                            if (i == 0)
                            {
                                Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V1");
                            }
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
                }
            else
            {
                errormessage = retval;
            }
            GC.Collect();
        }
    }
}
    
