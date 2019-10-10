using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.BusquedaReimpresiones.Estructura;
using CajaIndu.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura;


namespace CajaIndu.AppPersistencia.Class.NotasDeCredito
{
    class NotasDeCredito
    {
        public List<T_DOCUMENTOS> ObjDatosNC = new List<T_DOCUMENTOS>();
        public List<ESTADO> Retorno = new List<ESTADO>();
        public List<VIAS_PAGO2> ViasPago = new List<VIAS_PAGO2>();
        public string errormessage = "";
        public string protesto = "";
        ConexSAP connectorSap = new ConexSAP();

        public void notasdecredito(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_DOCUMENTO, string P_RUT,
    string P_SOCIEDAD, string P_LAND, string TipoBusqueda, string IDCAJA)
        {
            ObjDatosNC.Clear();
            Retorno.Clear();
            ViasPago.Clear();
            errormessage = "";
            protesto = "";
            IRfcTable lt_t_documentos;
            IRfcTable lt_viaspago;
            IRfcTable lt_retorno;
            T_DOCUMENTOS NC_resp;
            ESTADO retorno_resp;
            VIAS_PAGO2 VIASPAGO_resp;

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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_DOC_NC");
                BapiGetUser.SetValue("DOCUMENTO", P_DOCUMENTO);
                BapiGetUser.SetValue("LAND", P_LAND);
                BapiGetUser.SetValue("RUT", P_RUT);
                BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                BapiGetUser.SetValue("ID_CAJA", IDCAJA);
                BapiGetUser.Invoke(SapRfcDestination);

                lt_t_documentos = BapiGetUser.GetTable("T_DOCUMENTOS");
                lt_retorno = BapiGetUser.GetTable("RETORNO");
                lt_viaspago = BapiGetUser.GetTable("VIAS_PAGO");
                //lt_PART_ABIERTAS = BapiGetUser.GetTable("ZCLSP_TT_LISTA_DOCUMENTOS");
                try
                {
                    if (lt_t_documentos.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_t_documentos.RowCount; i++)
                        {
                            try
                            {
                                lt_t_documentos.CurrentIndex = i;
                                NC_resp = new T_DOCUMENTOS();

                                NC_resp.NDOCTO = lt_t_documentos[i].GetString("NDOCTO");
                                string Monto = "";
                                int indice = 0;
                                //*******
                                //if (lt_t_documentos[i].GetString("MONTOF") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO").Substring(0, indice - 1);
                                //    NC_resp.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    NC_resp.MONTOF = lt_t_documentos[i].GetString("MONTOF");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTOF").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    NC_resp.MONTOF = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    NC_resp.MONTOF = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_t_documentos[i].GetString("MONTO") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO").Substring(0, indice - 1);
                                //    NC_resp.MONTO = Monto;
                                //}
                                //else
                                //{
                                //    NC_resp.MONTO = lt_t_documentos[i].GetString("MONTO");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTO").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    NC_resp.MONTO = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTO"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    NC_resp.MONTO = string.Format("{0:0,0.##}", ValorAux);
                                }
                                NC_resp.MONEDA = lt_t_documentos[i].GetString("MONEDA");
                                NC_resp.FECVENCI = lt_t_documentos[i].GetString("FECVENCI");
                                NC_resp.CONTROL_CREDITO = lt_t_documentos[i].GetString("CONTROL_CREDITO");
                                NC_resp.CEBE = lt_t_documentos[i].GetString("CEBE");
                                NC_resp.COND_PAGO = lt_t_documentos[i].GetString("COND_PAGO");
                                NC_resp.RUTCLI = lt_t_documentos[i].GetString("RUTCLI");
                                NC_resp.NOMCLI = lt_t_documentos[i].GetString("NOMCLI");
                                NC_resp.ESTADO = lt_t_documentos[i].GetString("ESTADO");
                                NC_resp.ICONO = lt_t_documentos[i].GetString("ICONO");
                                NC_resp.DIAS_ATRASO = lt_t_documentos[i].GetString("DIAS_ATRASO");
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTO_ABONADO").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    NC_resp.MONTO_ABONADO = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTO_ABONADO"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    NC_resp.MONTO_ABONADO = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_t_documentos[i].GetString("MONTOF_ABON") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO_ABONADO").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO_ABONADO").Substring(0, indice - 1);
                                //    NC_resp.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    NC_resp.MONTOF_ABON = lt_t_documentos[i].GetString("MONTOF_ABON");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTOF_ABON").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    NC_resp.MONTOF_ABON = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF_ABON"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    NC_resp.MONTOF_ABON = string.Format("{0:0,0.##}", ValorAux);
                                }
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTO_PAGAR").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    NC_resp.MONTO_PAGAR = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTO_PAGAR"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    NC_resp.MONTO_PAGAR = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_t_documentos[i].GetString("MONTOF_PAGAR") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO_PAGAR").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO_PAGAR").Substring(0, indice - 1);
                                //    NC_resp.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    NC_resp.MONTOF_PAGAR = lt_t_documentos[i].GetString("MONTOF_PAGAR");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTOF_PAGAR").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    NC_resp.MONTOF_PAGAR = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF_PAGAR"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    NC_resp.MONTOF_PAGAR = string.Format("{0:0,0.##}", ValorAux);
                                }
                                NC_resp.NREF = lt_t_documentos[i].GetString("NREF");
                                NC_resp.FECHA_DOC = lt_t_documentos[i].GetString("FECHA_DOC");
                                NC_resp.COD_CLIENTE = lt_t_documentos[i].GetString("COD_CLIENTE");
                                NC_resp.SOCIEDAD = lt_t_documentos[i].GetString("SOCIEDAD");
                                NC_resp.CLASE_DOC = lt_t_documentos[i].GetString("CLASE_DOC");
                                NC_resp.CLASE_CUENTA = lt_t_documentos[i].GetString("CLASE_CUENTA");
                                NC_resp.CME = lt_t_documentos[i].GetString("CME");
                                NC_resp.ACC = lt_t_documentos[i].GetString("ACC");
                                NC_resp.FACT_SD_ORIGEN = lt_t_documentos[i].GetString("FACT_SD_ORIGEN");
                                NC_resp.FACT_ELECT = lt_t_documentos[i].GetString("FACT_ELECT");
                                NC_resp.ID_COMPROBANTE = lt_t_documentos[i].GetString("ID_COMPROBANTE");
                                NC_resp.ID_CAJA = lt_t_documentos[i].GetString("ID_CAJA");
                                NC_resp.LAND = lt_t_documentos[i].GetString("LAND");
                                NC_resp.BAPI = lt_t_documentos[i].GetString("BAPI");
                                ObjDatosNC.Add(NC_resp);
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
                       System.Windows.Forms.MessageBox.Show("No existe(n) registro(s)");
                    }

                    String Mensaje = "";
                    if (lt_retorno.Count > 0)
                    {
                        retorno_resp = new ESTADO();
                        for (int i = 0; i < lt_retorno.Count(); i++)
                        {
                            lt_retorno.CurrentIndex = i;

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
                     //   System.Windows.MessageBox.Show(Mensaje);
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
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    VIASPAGO_resp.MONTO = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_viaspago[i].GetString("MONTO"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    VIASPAGO_resp.MONTO =  string.Format("{0:0,0.##}", ValorAux);
                                }


                               // VIASPAGO_resp.MONTO = Convert.ToDouble(lt_viaspago[i].GetString("MONTO"));
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
                                ViasPago.Add(VIASPAGO_resp);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                            }
                        }
                    }
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                }

            }
            else
            {
                errormessage = retval;
                GC.Collect();
            }
        }

    }
}
