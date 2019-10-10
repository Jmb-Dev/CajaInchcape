using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura;
using CajaIndigo.AppPersistencia.Class.NotasDeCreditoCheck.Estructura;
using CajaIndigo.AppPersistencia.Class.BusquedaReimpresiones.Estructura;

namespace CajaIndigo.AppPersistencia.Class.NotasDeCreditoCheck
{
    class NotasDeCreditoCheck
    {
            //public List<T_DOCUMENTOS> ObjDatosMonitor = new List<T_DOCUMENTOS>();
            public string errormessage = "";
            public string message = "";
            public string IdCaja = "";
            public string Efectivo = "";

            public List<T_DOCUMENTOS> documentos = new List<T_DOCUMENTOS>();
            //DETALLE_VP
            public List<VIAS_PAGO2> viapago = new List<VIAS_PAGO2>();
            //RESUMEN_VP
            public List<DET_EFECT> det_efectivo = new List<DET_EFECT>();
            //DETALLE_REND
            //public List<DETALLE_REND> detalle_rend = new List<DETALLE_REND>();
            ConexSAP connectorSap = new ConexSAP();
            public List<RETURN2> T_Retorno = new List<RETURN2>();

            public void chequearnotascreditos(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_MONEDA, string P_PAIS, List<T_DOCUMENTOS> P_DOCSAPAGAR)
            {
                try
                {
                    RETURN2 retorno;
                    DET_EFECT efectivo;
                    T_DOCUMENTOS docs;
                    VIAS_PAGO2 vp;
                    //DETALLE_REND detallerend;
                    T_Retorno.Clear();
                    documentos.Clear();
                    viapago.Clear();
                    det_efectivo.Clear();
                    errormessage = "";
                    message = "";
                    IdCaja = "";
                    Efectivo = "0";

                    IRfcTable ls_RETORNO;
                    IRfcTable lt_VP;
                    IRfcTable lt_DOCS;
                    IRfcTable lt_EFECTIVO;

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

                        IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_CHECK_NC");

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
                                GralDat2.SetValue("ICONO", P_DOCSAPAGAR[i].ICONO);
                                GralDat2.SetValue("DIAS_ATRASO", P_DOCSAPAGAR[i].DIAS_ATRASO);
                                GralDat2.SetValue("MONTO_ABONADO", P_DOCSAPAGAR[i].MONTO_ABONADO);
                                GralDat2.SetValue("MONTOF_ABON", P_DOCSAPAGAR[i].MONTOF_ABON);
                                GralDat2.SetValue("MONTO_PAGAR", P_DOCSAPAGAR[i].MONTO_PAGAR);
                                GralDat2.SetValue("MONTOF_PAGAR", P_DOCSAPAGAR[i].MONTOF_PAGAR);
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
                                    docs = new T_DOCUMENTOS ();

                                    docs.NDOCTO = lt_DOCS[i].GetString("NDOCTO");
                                    string Monto = "";
                                    int indice = 0;
                                    //*******
                                    if (lt_DOCS[i].GetString("MONTOF") == "")
                                    {
                                        indice = lt_DOCS[i].GetString("MONTO").IndexOf(',');
                                        Monto = lt_DOCS[i].GetString("MONTO").Substring(0, indice - 1);
                                        docs.MONTOF = Monto;
                                    }
                                    else
                                    {
                                        docs.MONTOF = lt_DOCS[i].GetString("MONTOF");
                                    }
                                    if (lt_DOCS[i].GetString("MONTO") == "")
                                    {
                                        indice = lt_DOCS[i].GetString("MONTO").IndexOf(',');
                                        Monto = lt_DOCS[i].GetString("MONTO").Substring(0, indice - 1);
                                        docs.MONTO = Monto;
                                    }
                                    else
                                    {
                                        docs.MONTO = lt_DOCS[i].GetString("MONTO");
                                    }
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
                                    if (lt_DOCS[i].GetString("MONTOF_ABON") == "")
                                    {
                                        indice = lt_DOCS[i].GetString("MONTO_ABONADO").IndexOf(',');
                                        Monto = lt_DOCS[i].GetString("MONTO_ABONADO").Substring(0, indice - 1);
                                        docs.MONTOF = Monto;
                                    }
                                    else
                                    {
                                        docs.MONTOF_ABON = lt_DOCS[i].GetString("MONTOF_ABON");
                                    }
                                    if (lt_DOCS[i].GetString("MONTOF_PAGAR") == "")
                                    {
                                        indice = lt_DOCS[i].GetString("MONTO_PAGAR").IndexOf(',');
                                        Monto = lt_DOCS[i].GetString("MONTO_PAGAR").Substring(0, indice - 1);
                                        docs.MONTOF = Monto;
                                    }
                                    else
                                    {
                                        docs.MONTOF_PAGAR = lt_DOCS[i].GetString("MONTOF_PAGAR");
                                    }
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

                        lt_EFECTIVO = BapiGetUser.GetTable("DET_EFECT");
                        try
                        {
                            for (int i = 0; i < lt_EFECTIVO.Count(); i++)
                            {
                                lt_EFECTIVO.CurrentIndex = i;
                                efectivo = new DET_EFECT();
                                efectivo.LAND = lt_EFECTIVO.GetString("LAND");
                                efectivo.ID_CAJA = lt_EFECTIVO.GetString("ID_CAJA");
                                efectivo.SOCIEDAD = lt_EFECTIVO.GetString("SOCIEDAD");
                                efectivo.SOCIEDAD_TXT = lt_EFECTIVO.GetString("SOCIEDAD_TXT");
                                efectivo.VIA_PAGO = lt_EFECTIVO.GetString("VIA_PAGO");
                                efectivo.TEXT1 = lt_EFECTIVO.GetString("TEXT1");
                                efectivo.MONEDA = lt_EFECTIVO.GetString("MONEDA");
                                efectivo.MONTO = lt_EFECTIVO.GetString("MONTO");
                                Efectivo = Convert.ToString(Convert.ToDouble(Efectivo) + Convert.ToDouble(lt_EFECTIVO.GetString("MONTO")));
                                efectivo.CANT_DOCS = lt_EFECTIVO.GetString("CANT_DOCS");
                                det_efectivo.Add(efectivo);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message, ex.StackTrace);
                            MessageBox.Show(ex.Message + ex.StackTrace);
                        }
                        // PART_ABIERTAS_resp.MONTOF_ABON = string.Format("{0:0.##}", Cualquiernombre);
                    //    Efectivo = string.Format("{0:0,0}", Efectivo);
                        //Efectivo = Efectivo.ToString("0.0", CultureInfo.InvariantCulture);
                        //Efectivo = Efectivo.ToString("0:0,0");
                        //string Efect = string.Format("{0:0,0}", Efectivo);
                        //double value = 1234567890;
                        //Console.WriteLine(value.ToString("#,#", CultureInfo.InvariantCulture));
                        //Console.WriteLine(String.Format(CultureInfo.InvariantCulture,
                        //                                "{0:0,#}", value));
                        Double Efect = Convert.ToDouble(Efectivo);
                        CultureInfo elGR = CultureInfo.CreateSpecificCulture("el-GR");
                        Console.WriteLine(Efect.ToString("0,0", elGR));
                        Console.WriteLine(String.Format(elGR, "{0:0,0}", Efect));
                        Efectivo = Efect.ToString("0,0", elGR);
                        Efectivo = String.Format(elGR, "{0:0,0}", Efect);
                       // Efectivo = Convert.ToString(Efect).Replace(",", "");
                       
                
                        lt_VP = BapiGetUser.GetTable("VIAS_PAGO");
                        if (lt_VP.Count > 0)
                        {
                            //LLenamos la tabla de salida lt_DatGen
                            for (int i = 0; i < lt_VP.RowCount; i++)
                            {
                                try
                                {
                                    lt_VP.CurrentIndex = i;
                                    vp = new VIAS_PAGO2();

                                    vp.MANDT = lt_VP[i].GetString("MANDT");
                                    vp.LAND = lt_VP[i].GetString("LAND");
                                    vp.ID_COMPROBANTE = lt_VP[i].GetString("ID_COMPROBANTE");
                                    vp.ID_DETALLE = lt_VP[i].GetString("ID_DETALLE");
                                    vp.VIA_PAGO = lt_VP[i].GetString("VIA_PAGO");
                                    vp.MONTO = lt_VP[i].GetString("MONTO");
                                    vp.MONEDA = lt_VP[i].GetString("MONEDA");
                                    vp.BANCO = lt_VP[i].GetString("BANCO");
                                    vp.EMISOR = lt_VP[i].GetString("EMISOR");
                                    vp.NUM_CHEQUE = lt_VP[i].GetString("NUM_CHEQUE");
                                    vp.COD_AUTORIZACION = lt_VP[i].GetString("COD_AUTORIZACION");
                                    vp.NUM_CUOTAS = lt_VP[i].GetString("NUM_CUOTAS");
                                    vp.FECHA_VENC = lt_VP[i].GetString("FECHA_VENC");
                                    vp.TEXTO_POSICION = lt_VP[i].GetString("TEXTO_POSICION");
                                    vp.ANEXO = lt_VP[i].GetString("ANEXO");
                                    vp.SUCURSAL = lt_VP[i].GetString("SUCURSAL");
                                    vp.NUM_CUENTA = lt_VP[i].GetString("NUM_CUENTA");
                                    vp.NUM_TARJETA = lt_VP[i].GetString("NUM_TARJETA");
                                    vp.NUM_VALE_VISTA = lt_VP[i].GetString("NUM_VALE_VISTA");
                                    vp.PATENTE = lt_VP[i].GetString("PATENTE");
                                    vp.NUM_VENTA = lt_VP[i].GetString("NUM_VENTA");
                                    vp.PAGARE = lt_VP[i].GetString("PAGARE");
                                    vp.FECHA_EMISION = lt_VP[i].GetString("FECHA_EMISION");
                                    vp.NOMBRE_GIRADOR = lt_VP[i].GetString("NOMBRE_GIRADOR");
                                    vp.CARTA_CURSE = lt_VP[i].GetString("CARTA_CURSE");
                                    vp.NUM_TRANSFER = lt_VP[i].GetString("NUM_TRANSFER");
                                    vp.NUM_DEPOSITO = lt_VP[i].GetString("NUM_DEPOSITO");
                                    vp.CTA_BANCO = lt_VP[i].GetString("CTA_BANCO");
                                    vp.IFINAN = lt_VP[i].GetString("IFINAN");
                                    vp.CORRE = lt_VP[i].GetString("CORRE");
                                    vp.ZUONR = lt_VP[i].GetString("ZUONR");
                                    vp.HKONT = lt_VP[i].GetString("HKONT");
                                    vp.PRCTR = lt_VP[i].GetString("PRCTR");
                                    vp.ZNOP = lt_VP[i].GetString("ZNOP");
                                    viapago.Add(vp);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message + ex.StackTrace);
                                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                                }
                            }
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
                                    message = message + "-" + ls_RETORNO.GetString("MESSAGE") + ":" + ls_RETORNO.GetString("MESSAGE_V1").Trim(); ;
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
                       // errormessage = retval;
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

