using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CajaIndu.AppPersistencia.Class.BusquedaAnulacion.Estructura;
using CajaIndu.AppPersistencia.Class.BusquedaReimpresiones.Estructura;
using CajaIndu.AppPersistencia.Class.Connections;
using SAP.Middleware.Connector;


namespace CajaIndu.AppPersistencia.Class.BusquedaReimpresiones
{
    class BusquedaReimpresiones
    {
        public List<DOCUMENTOS> Documentos = new List<DOCUMENTOS>();
        public List<VIAS_PAGO2> ViasPago = new List<VIAS_PAGO2>();
        public List<RETORNO> Retorno = new List<RETORNO>();
        public string errormessage = "";
        public string message = "";
        ConexSAP connectorSap = new ConexSAP();

        public void docsreimpresion(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE
            , string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_COMPROBANTE, string P_RUT
            , string P_ID_APERTURA, string P_LAND, string P_IDCAJA, string P_BATCH)
        {
            Documentos.Clear();
            ViasPago.Clear();
            Retorno.Clear();
            errormessage = "";
            message = "";
            IRfcTable lt_documentos;
            IRfcTable lt_viaspago;
            IRfcTable lt_retorno;

            //  PART_ABIERTAS  PART_ABIERTAS_resp;
            DOCUMENTOS DOCUMENTOS_resp;
            VIAS_PAGO2 VIASPAGO_resp;
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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_BUSCA_COMP_REIMP");
                BapiGetUser.SetValue("ID_COMPROBANTE", P_COMPROBANTE);
                BapiGetUser.SetValue("RUT", P_RUT);
                BapiGetUser.SetValue("LAND", P_LAND);
                BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);
                BapiGetUser.SetValue("ID_CAJA", P_IDCAJA);
                BapiGetUser.SetValue("BATCH", P_BATCH);
                BapiGetUser.Invoke(SapRfcDestination);

                lt_documentos = BapiGetUser.GetTable("DOCUMENTOS");
                lt_viaspago = BapiGetUser.GetTable("VIAS_PAGO");
                lt_retorno = BapiGetUser.GetTable("RETORNO");

                if (lt_documentos.Count > 0)
                {
                    //LLenamos la tabla de salida lt_DatGen
                    for (int i = 0; i < lt_documentos.RowCount; i++)
                    {
                        try
                        {
                            lt_documentos.CurrentIndex = i;
                            DOCUMENTOS_resp = new DOCUMENTOS();

                            DOCUMENTOS_resp.MANDT = lt_documentos[i].GetString("MANDT");
                            DOCUMENTOS_resp.LAND = lt_documentos[i].GetString("LAND");
                            DOCUMENTOS_resp.ID_COMPROBANTE = lt_documentos[i].GetString("ID_COMPROBANTE");
                            DOCUMENTOS_resp.POSICION = lt_documentos[i].GetString("POSICION");
                            DOCUMENTOS_resp.CLIENTE = lt_documentos[i].GetString("CLIENTE");
                            DOCUMENTOS_resp.TIPO_DOCUMENTO = lt_documentos[i].GetString("TIPO_DOCUMENTO");
                            DOCUMENTOS_resp.SOCIEDAD = lt_documentos[i].GetString("SOCIEDAD");
                            DOCUMENTOS_resp.NRO_DOCUMENTO = lt_documentos[i].GetString("NRO_DOCUMENTO");
                            DOCUMENTOS_resp.NRO_REFERENCIA = lt_documentos[i].GetString("NRO_REFERENCIA");
                            DOCUMENTOS_resp.CAJERO_RESP = lt_documentos[i].GetString("CAJERO_RESP");
                            DOCUMENTOS_resp.CAJERO_GEN = lt_documentos[i].GetString("CAJERO_GEN");
                            DOCUMENTOS_resp.ID_CAJA = lt_documentos[i].GetString("ID_CAJA");
                            DOCUMENTOS_resp.FECHA_COMP = lt_documentos[i].GetString("FECHA_COMP");
                            DOCUMENTOS_resp.HORA = lt_documentos[i].GetString("HORA");
                            DOCUMENTOS_resp.NRO_COMPENSACION = lt_documentos[i].GetString("NRO_COMPENSACION");
                            DOCUMENTOS_resp.TEXTO_CABECERA = lt_documentos[i].GetString("TEXTO_CABECERA");
                            DOCUMENTOS_resp.NULO = lt_documentos[i].GetString("NULO");
                            DOCUMENTOS_resp.USR_ANULADOR = lt_documentos[i].GetString("USR_ANULADOR");
                            DOCUMENTOS_resp.NRO_ANULACION = lt_documentos[i].GetString("NRO_ANULACION");
                            DOCUMENTOS_resp.APROBADOR_ANULA = lt_documentos[i].GetString("APROBADOR_ANULA");
                            DOCUMENTOS_resp.TXT_ANULACION = lt_documentos[i].GetString("TXT_ANULACION");
                            DOCUMENTOS_resp.EXCEPCION = lt_documentos[i].GetString("EXCEPCION");
                            DOCUMENTOS_resp.FECHA_DOC = lt_documentos[i].GetString("FECHA_DOC");
                            DOCUMENTOS_resp.FECHA_VENC_DOC = lt_documentos[i].GetString("FECHA_VENC_DOC");
                            DOCUMENTOS_resp.NUM_CUOTA = lt_documentos[i].GetString("NUM_CUOTA");
                            if (lt_documentos[i].GetString("MONEDA") == "CLP")
                            {
                                string Valor = lt_documentos[i].GetString("MONTO_DOC").Trim();
                                if (Valor.Contains("-"))
                                {
                                    Valor = "-" + Valor.Replace("-", "");
                                }
                                Valor = Valor.Replace(".", "");
                                Valor = Valor.Replace(",", "");
                                decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length-2));
                                string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                DOCUMENTOS_resp.MONTO_DOC = Cualquiernombre;
                            }
                            else
                            {
                                string moneda = Convert.ToString(lt_documentos[i].GetString("MONTO_DOC"));
                                decimal ValorAux = Convert.ToDecimal(moneda);
                                DOCUMENTOS_resp.MONTO_DOC = string.Format("{0:0,0.##}", ValorAux);
                            }
                            //DOCUMENTOS_resp.MONTO_DOC = lt_documentos[i].GetString("MONTO_DOC");
                            if (lt_documentos[i].GetString("MONEDA") == "CLP")
                            {
                                string Valor = lt_documentos[i].GetString("MONTO_DIFERENCIA").Trim();
                                if (Valor.Contains("-"))
                                {
                                    Valor = "-" + Valor.Replace("-", "");
                                }
                                Valor = Valor.Replace(".", "");
                                Valor = Valor.Replace(",", "");
                                decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                DOCUMENTOS_resp.MONTO_DIFERENCIA = Cualquiernombre;
                            }
                            else
                            {
                                string moneda = Convert.ToString(lt_documentos[i].GetString("MONTO_DIFERENCIA"));
                                decimal ValorAux = Convert.ToDecimal(moneda);
                                DOCUMENTOS_resp.MONTO_DIFERENCIA = string.Format("{0:0,0.##}", ValorAux);
                            }
                            //DOCUMENTOS_resp.MONTO_DIFERENCIA = lt_documentos[i].GetString("MONTO_DIFERENCIA");
                            DOCUMENTOS_resp.TEXTO_EXCEPCION = lt_documentos[i].GetString("TEXTO_EXCEPCION");
                            DOCUMENTOS_resp.PARCIAL = lt_documentos[i].GetString("PARCIAL");
                            DOCUMENTOS_resp.TIME = lt_documentos[i].GetString("TIME");
                            DOCUMENTOS_resp.APROBADOR_EX = lt_documentos[i].GetString("APROBADOR_EX");
                            DOCUMENTOS_resp.MONEDA = lt_documentos[i].GetString("MONEDA");
                            DOCUMENTOS_resp.CLASE_CUENTA = lt_documentos[i].GetString("CLASE_CUENTA");
                            DOCUMENTOS_resp.CLASE_DOC = lt_documentos[i].GetString("CLASE_DOC");
                            DOCUMENTOS_resp.NUM_CANCELACION = lt_documentos[i].GetString("NUM_CANCELACION");
                            DOCUMENTOS_resp.CME = lt_documentos[i].GetString("CME");
                            DOCUMENTOS_resp.NOTA_VENTA = lt_documentos[i].GetString("NOTA_VENTA");
                            DOCUMENTOS_resp.CEBE = lt_documentos[i].GetString("CEBE");
                            DOCUMENTOS_resp.ACC = lt_documentos[i].GetString("ACC");
                            Documentos.Add(DOCUMENTOS_resp);
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

                if (lt_viaspago.Count > 0)
                {
                    //LLenamos la tabla de salida lt_DatGen
                    for (int i = 0; i < lt_viaspago.RowCount; i++)
                    {
                        try
                        {
                            lt_viaspago.CurrentIndex = i;
                            VIASPAGO_resp = new VIAS_PAGO2();

                            VIASPAGO_resp.MANDT = lt_viaspago[i].GetString("MANDT");
                            VIASPAGO_resp.LAND = lt_viaspago[i].GetString("LAND");
                            VIASPAGO_resp.ID_COMPROBANTE = lt_viaspago[i].GetString("ID_COMPROBANTE");
                            VIASPAGO_resp.ID_DETALLE = lt_viaspago[i].GetString("ID_DETALLE");
                            VIASPAGO_resp.VIA_PAGO = lt_viaspago[i].GetString("VIA_PAGO");
                            if (lt_viaspago[i].GetString("MONEDA") == "CLP")
                            {
                                string Valor = lt_viaspago[i].GetString("MONTO").Trim();
                                if (Valor.Contains("-"))
                                {
                                    Valor = "-" + Valor.Replace("-", "");
                                }
                                Valor = Valor.Replace(".", "");
                                Valor = Valor.Replace(",", "");
                                //decimal ValorAux = Convert.ToDecimal(Valor);
                                decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                VIASPAGO_resp.MONTO = Cualquiernombre;
                            }
                            else
                            {
                                string moneda = Convert.ToString(lt_viaspago[i].GetString("MONTO"));
                                decimal ValorAux = Convert.ToDecimal(moneda);
                                VIASPAGO_resp.MONTO = string.Format("{0:0,0.##}", ValorAux);
                            }
                            //VIASPAGO_resp.MONTO =  lt_viaspago[i].GetString("MONTO");
                            VIASPAGO_resp.MONEDA = lt_viaspago[i].GetString("MONEDA");
                            VIASPAGO_resp.BANCO = lt_viaspago[i].GetString("BANCO");
                            VIASPAGO_resp.EMISOR = lt_viaspago[i].GetString("EMISOR");
                            VIASPAGO_resp.NUM_CHEQUE = lt_viaspago[i].GetString("NUM_CHEQUE");
                            VIASPAGO_resp.COD_AUTORIZACION = lt_viaspago[i].GetString("COD_AUTORIZACION");
                            VIASPAGO_resp.NUM_CUOTAS = lt_viaspago[i].GetString("NUM_CUOTAS");
                            VIASPAGO_resp.FECHA_VENC = lt_viaspago[i].GetString("FECHA_VENC");
                            VIASPAGO_resp.TEXTO_POSICION = lt_viaspago[i].GetString("TEXTO_POSICION");
                            VIASPAGO_resp.ANEXO = lt_viaspago[i].GetString("ANEXO");
                            VIASPAGO_resp.SUCURSAL = lt_viaspago[i].GetString("SUCURSAL");
                            VIASPAGO_resp.NUM_CUENTA = lt_viaspago[i].GetString("NUM_CUENTA");
                            VIASPAGO_resp.NUM_TARJETA = lt_viaspago[i].GetString("NUM_TARJETA");
                            VIASPAGO_resp.NUM_VALE_VISTA = lt_viaspago[i].GetString("NUM_VALE_VISTA");
                            VIASPAGO_resp.PATENTE = lt_viaspago[i].GetString("PATENTE");
                            VIASPAGO_resp.NUM_VENTA = lt_viaspago[i].GetString("NUM_VENTA");
                            VIASPAGO_resp.PAGARE = lt_viaspago[i].GetString("PAGARE");
                            VIASPAGO_resp.FECHA_EMISION = lt_viaspago[i].GetString("FECHA_EMISION");
                            VIASPAGO_resp.NOMBRE_GIRADOR = lt_viaspago[i].GetString("NOMBRE_GIRADOR");
                            VIASPAGO_resp.CARTA_CURSE = lt_viaspago[i].GetString("CARTA_CURSE");
                            VIASPAGO_resp.NUM_TRANSFER = lt_viaspago[i].GetString("NUM_TRANSFER");
                            VIASPAGO_resp.NUM_DEPOSITO = lt_viaspago[i].GetString("NUM_DEPOSITO");
                            VIASPAGO_resp.CTA_BANCO = lt_viaspago[i].GetString("CTA_BANCO");
                            VIASPAGO_resp.IFINAN = lt_viaspago[i].GetString("IFINAN");
                            VIASPAGO_resp.CORRE = lt_viaspago[i].GetString("CORRE");
                            VIASPAGO_resp.ZUONR = lt_viaspago[i].GetString("ZUONR");
                            VIASPAGO_resp.HKONT = lt_viaspago[i].GetString("HKONT");
                            VIASPAGO_resp.PRCTR = lt_viaspago[i].GetString("PRCTR");
                            VIASPAGO_resp.ZNOP = lt_viaspago[i].GetString("ZNOP");
                            ViasPago.Add(VIASPAGO_resp);
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
                            Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE");
                            retorno_resp.MESSAGE_V2 = lt_retorno.GetString("MESSAGE_V2");
                            retorno_resp.MESSAGE_V3 = lt_retorno.GetString("MESSAGE_V3");
                            retorno_resp.MESSAGE_V4 = lt_retorno.GetString("MESSAGE_V4");
                            retorno_resp.PARAMETER = lt_retorno.GetString("PARAMETER");
                            retorno_resp.ROW = lt_retorno.GetString("ROW");
                            retorno_resp.FIELD = lt_retorno.GetString("FIELD");
                            retorno_resp.SYSTEM = lt_retorno.GetString("SYSTEM");
                            Retorno.Add(retorno_resp);
                        }
                        //System.Windows.MessageBox.Show(Mensaje);
                    }
                    //else
                    //{
                    //    System.Windows.MessageBox.Show("No existe(n) registro(s)");
                    //}
           

                    GC.Collect();
        }
        else
        {
            errormessage = retval;
        }
            
           
            GC.Collect();
        }
    }
}
    


    