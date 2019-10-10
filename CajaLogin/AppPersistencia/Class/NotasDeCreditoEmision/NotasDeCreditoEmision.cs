using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura;
using CajaIndigo.AppPersistencia.Class.NotasDeCreditoCheck.Estructura;

namespace CajaIndigo.AppPersistencia.Class.NotasDeCreditoEmision
{
    class NotasDeCreditoEmision
    {
        public string errormessage = "";
        public string message = "";
        public string IdCaja = "";
        public string Efectivo = "";
        public string NumComprob = "";

        public List<T_DOCUMENTOS> documentos = new List<T_DOCUMENTOS>();
        ConexSAP connectorSap = new ConexSAP();
        public List<RETURN2> T_Retorno = new List<RETURN2>();

        public void emitirnotasdecredito(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE
            , string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_MONEDA, string P_PAIS
            , List<T_DOCUMENTOS> P_DOCSAPAGAR, string P_CHKTRIB)
        {
            try
            {
                RETURN2 retorno;
                
                T_DOCUMENTOS docs;
                
                //DETALLE_REND detallerend;
                T_Retorno.Clear();
                documentos.Clear();
                errormessage = "";
                message = "";
                IdCaja = "";
                Efectivo = "0";

                IRfcTable ls_RETORNO;
                IRfcTable lt_DOCS;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_REC_Y_FAC_NC");

                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("PAY_CURRENCY", P_MONEDA);
                    BapiGetUser.SetValue("LAND", P_PAIS);
                    IRfcTable GralDat2 = BapiGetUser.GetTable("DOCUMENTOS");
                    try
                    {
                        for (var i = 0; i < P_DOCSAPAGAR.Count; i++)
                        {
                            GralDat2.Append();
                            GralDat2.SetValue("NDOCTO", P_DOCSAPAGAR[i].NDOCTO);
                            GralDat2.SetValue("MONTO", P_DOCSAPAGAR[i].MONTO);
                            GralDat2.SetValue("MONTOF", P_DOCSAPAGAR[i].MONTOF);
                            GralDat2.SetValue("MONEDA", P_DOCSAPAGAR[i].MONEDA);
                            GralDat2.SetValue("LAND", P_DOCSAPAGAR[i].LAND);
                            GralDat2.SetValue("FECVENCI", P_DOCSAPAGAR[i].FECVENCI);
                            GralDat2.SetValue("CONTROL_CREDITO", P_DOCSAPAGAR[i].CONTROL_CREDITO);
                            GralDat2.SetValue("CEBE", P_DOCSAPAGAR[i].CEBE);
                            GralDat2.SetValue("COND_PAGO", P_DOCSAPAGAR[i].COND_PAGO);
                            GralDat2.SetValue("RUTCLI", P_DOCSAPAGAR[i].RUTCLI);
                            GralDat2.SetValue("NOMCLI", P_DOCSAPAGAR[i].NOMCLI);
                            GralDat2.SetValue("ESTADO", P_DOCSAPAGAR[i].ESTADO);
                            GralDat2.SetValue("ICONO",P_CHKTRIB);
                            GralDat2.SetValue("DIAS_ATRASO", P_DOCSAPAGAR[i].DIAS_ATRASO);
                            string MontoAux = "";
                            MontoAux = P_DOCSAPAGAR[i].MONTO_ABONADO;
                            MontoAux = MontoAux.Replace(",", "");
                            MontoAux = MontoAux.Replace(".", "");
                            GralDat2.SetValue("MONTO_ABONADO", MontoAux );
                            MontoAux = P_DOCSAPAGAR[i].MONTOF_ABON;
                            MontoAux = MontoAux.Replace(",", "");
                            MontoAux = MontoAux.Replace(".", "");
                            GralDat2.SetValue("MONTOF_ABON", MontoAux);
                            MontoAux = P_DOCSAPAGAR[i].MONTO_PAGAR;
                            MontoAux = MontoAux.Replace(",", "");
                            MontoAux = MontoAux.Replace(".", "");
                            GralDat2.SetValue("MONTO_PAGAR", MontoAux);
                            MontoAux = P_DOCSAPAGAR[i].MONTOF_PAGAR;
                            MontoAux = MontoAux.Replace(",", "");
                            MontoAux = MontoAux.Replace(".", "");
                            GralDat2.SetValue("MONTOF_PAGAR", MontoAux);
                            GralDat2.SetValue("NREF", P_DOCSAPAGAR[i].NREF);
                            GralDat2.SetValue("FECHA_DOC", P_DOCSAPAGAR[i].FECHA_DOC);
                            GralDat2.SetValue("COD_CLIENTE", P_DOCSAPAGAR[i].COD_CLIENTE);
                            GralDat2.SetValue("SOCIEDAD", P_DOCSAPAGAR[i].SOCIEDAD);
                            GralDat2.SetValue("CLASE_DOC", P_DOCSAPAGAR[i].CLASE_DOC);
                            GralDat2.SetValue("CLASE_CUENTA", P_DOCSAPAGAR[i].CLASE_CUENTA);
                            GralDat2.SetValue("CME", P_DOCSAPAGAR[i].CME);
                            GralDat2.SetValue("ACC", P_DOCSAPAGAR[i].ACC);
                            GralDat2.SetValue("FACT_SD_ORIGEN", P_DOCSAPAGAR[i].FACT_SD_ORIGEN);
                            GralDat2.SetValue("FACT_ELECT", P_DOCSAPAGAR[i].FACT_ELECT);
                            GralDat2.SetValue("ID_COMPROBANTE", P_DOCSAPAGAR[i].ID_COMPROBANTE);
                            GralDat2.SetValue("ID_CAJA", P_DOCSAPAGAR[i].ID_CAJA);
                            GralDat2.SetValue("LAND", P_DOCSAPAGAR[i].LAND);
                            GralDat2.SetValue("BAPI", P_DOCSAPAGAR[i].BAPI);

                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                        //System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                    BapiGetUser.SetValue("DOCUMENTOS", GralDat2);


                    BapiGetUser.Invoke(SapRfcDestination);
                    //BapiGetUser.SetValue("I_VBELN",P_NUMDOCSD);
                    //IRfcTable GralDat = BapiGetUser.GetTable("VIAS_PAGO");
                    lt_DOCS = BapiGetUser.GetTable("DOCUMENTOS");

                    if (lt_DOCS.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_DOCS.RowCount; i++)
                        {
                            try
                            {
                                lt_DOCS.CurrentIndex = i;
                                docs = new T_DOCUMENTOS();

                                docs.NDOCTO = lt_DOCS[i].GetString("NDOCTO");
                                string Monto = "";
                                int indice = 0;
                                //*******
                                if (lt_DOCS[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_DOCS[i].GetString("MONTOF").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = 0;
                                    if ((Valor == "00" ) | (Valor == "0"))
                                    {
                                       ValorAux = Convert.ToDecimal(Valor);
                                    }
                                    else
                                    {
                                       ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    }
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    docs.MONTOF = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DOCS[i].GetString("MONTOF"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    docs.MONTOF = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_DOCS[i].GetString("MONTOF") == "")
                                //{
                                //    indice = lt_DOCS[i].GetString("MONTO").IndexOf(',');
                                //    Monto = lt_DOCS[i].GetString("MONTO").Substring(0, indice - 1);
                                //    docs.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    docs.MONTOF = lt_DOCS[i].GetString("MONTOF");
                                //}
                                if (lt_DOCS[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_DOCS[i].GetString("MONTO").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    if ((Valor == "00") | (Valor == "0"))
                                    {
                                        ValorAux = Convert.ToDecimal(Valor);
                                    }
                                    else
                                    {
                                        ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    }
                                    docs.MONTO = string.Format("{0:0,0}", ValorAux);
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DOCS[i].GetString("MONTO"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    docs.MONTO = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_DOCS[i].GetString("MONTO") == "")
                                //{
                                //    indice = lt_DOCS[i].GetString("MONTO").IndexOf(',');
                                //    Monto = lt_DOCS[i].GetString("MONTO").Substring(0, indice - 1);
                                //    docs.MONTO = Monto;
                                //}
                                //else
                                //{
                                //    docs.MONTO = lt_DOCS[i].GetString("MONTO");
                                //}
                                
                                docs.MONEDA = lt_DOCS[i].GetString("MONEDA");
                                docs.FECVENCI = lt_DOCS[i].GetString("FECVENCI");
                                docs.CONTROL_CREDITO = lt_DOCS[i].GetString("CONTROL_CREDITO");
                                docs.CEBE = lt_DOCS[i].GetString("CEBE");
                                docs.COND_PAGO = lt_DOCS[i].GetString("COND_PAGO");
                                docs.RUTCLI = lt_DOCS[i].GetString("RUTCLI");
                                docs.NOMCLI = lt_DOCS[i].GetString("NOMCLI");
                                docs.ESTADO = lt_DOCS[i].GetString("ESTADO");
                                docs.ICONO = lt_DOCS[i].GetString("ICONO");
                                docs.DIAS_ATRASO = lt_DOCS[i].GetString("DIAS_ATRASO");
                                if (lt_DOCS[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_DOCS[i].GetString("MONTOF_ABON").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux;
                                    if ((Valor == "00") | (Valor == "0"))
                                    {
                                        ValorAux = Convert.ToDecimal(Valor);
                                    }
                                    else
                                    {
                                        ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    }
                                    docs.MONTOF_ABON = string.Format("{0:0,0}", ValorAux);
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DOCS[i].GetString("MONTOF_ABON"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    docs.MONTOF_ABON = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_DOCS[i].GetString("MONTOF_ABON") == "")
                                //{
                                //    indice = lt_DOCS[i].GetString("MONTO_ABONADO").IndexOf(',');
                                //    Monto = lt_DOCS[i].GetString("MONTO_ABONADO").Substring(0, indice - 1);
                                //    docs.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    docs.MONTOF_ABON = lt_DOCS[i].GetString("MONTOF_ABON");
                                //}
                                if (lt_DOCS[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_DOCS[i].GetString("MONTOF_PAGAR").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    if ((Valor == "00") | (Valor == "0"))
                                    {
                                        ValorAux = Convert.ToDecimal(Valor);
                                    }
                                    else
                                    {
                                        ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    }
                                    docs.MONTOF_PAGAR = string.Format("{0:0,0}", ValorAux);
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DOCS[i].GetString("MONTOF_PAGAR"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    docs.MONTOF_PAGAR = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //if (lt_DOCS[i].GetString("MONTOF_PAGAR") == "")
                                //{
                                //    indice = lt_DOCS[i].GetString("MONTO_PAGAR").IndexOf(',');
                                //    Monto = lt_DOCS[i].GetString("MONTO_PAGAR").Substring(0, indice - 1);
                                //    docs.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    docs.MONTOF_PAGAR = lt_DOCS[i].GetString("MONTOF_PAGAR");
                                //}
                                docs.NREF = lt_DOCS[i].GetString("NREF");
                                docs.FECHA_DOC = lt_DOCS[i].GetString("FECHA_DOC");
                                docs.COD_CLIENTE = lt_DOCS[i].GetString("COD_CLIENTE");
                                docs.SOCIEDAD = lt_DOCS[i].GetString("SOCIEDAD");
                                docs.CLASE_DOC = lt_DOCS[i].GetString("CLASE_DOC");
                                docs.CLASE_CUENTA = lt_DOCS[i].GetString("CLASE_CUENTA");
                                docs.CME = lt_DOCS[i].GetString("CME");
                                docs.ACC = lt_DOCS[i].GetString("ACC");
                                docs.FACT_SD_ORIGEN = lt_DOCS[i].GetString("FACT_SD_ORIGEN");
                                docs.FACT_ELECT = lt_DOCS[i].GetString("FACT_ELECT");
                                docs.ID_COMPROBANTE = lt_DOCS[i].GetString("ID_COMPROBANTE");
                                docs.ID_CAJA = lt_DOCS[i].GetString("ID_CAJA");
                                docs.LAND = lt_DOCS[i].GetString("LAND");
                                documentos.Add(docs);
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

                    ls_RETORNO = BapiGetUser.GetTable("RETURN");
                    try
                    {
                        for (int i = 0; i < ls_RETORNO.Count(); i++)
                        {
                            ls_RETORNO.CurrentIndex = i;
                            retorno = new RETURN2();
                            if (ls_RETORNO.GetString("TYPE") == "S")
                            {
                                    message = message + "-" + ls_RETORNO.GetString("MESSAGE") + ":" + ls_RETORNO.GetString("MESSAGE_V1").Trim() + "\n"; 
                                    NumComprob = ls_RETORNO.GetString("MESSAGE_V4").Trim();
                            }
                            if (ls_RETORNO.GetString("TYPE") == "E")
                            {
                                errormessage = errormessage + " - " + ls_RETORNO.GetString("MESSAGE") + "\n";
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
                                // comprobante = ls_RETORNO.GetString("MESSAGE_V4");
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

