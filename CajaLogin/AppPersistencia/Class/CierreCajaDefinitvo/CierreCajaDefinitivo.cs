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
using CajaIndigo.AppPersistencia.Class.PreCierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.ArqueoCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCajaDefinitvo.Estructura;

namespace CajaIndigo.AppPersistencia.Class.CierreCajaDefinitvo
{
    class CierreCajaDefinitivo
    {
        public string errormessage = "";
        public string message = "";
        public string numerocierre = "";
        public string diasatraso = "";
        //Totalizadores por Via de Pago
        public double MontoIngresos = 0;
        public double MontoEfect = 0;
        public double MontoChqDia = 0;
        public double MontoChqFech = 0;
        public double MontoTransf = 0;
        public double MontoValeV = 0;
        public double MontoDepot = 0;
        public double MontoTarj = 0;
        public double MontoFinanc = 0;
        public double MontoApp = 0;
        public double MontoCredit = 0;
        public double MontoCCurse = 0;
        public double MontoEgresos = 0;
        public double MontoFondosFijos = 0;
        public double SaldoTotal = 0;
        //DETALLE_VP
        //public List<CAB_ARQUEO> cab_arqueo = new List<CAB_ARQUEO>();
        //RESUMEN_VP
        public List<DET_EFECTIVO> det_efectivo = new List<DET_EFECTIVO>();
        //RESUMEN_VP
        public List<RESUMEN_VP> resumen_viapago = new List<RESUMEN_VP>();
        //DETALLE_ARQUEO
        public List<DETALLE_ARQUEO> det_arqueo = new List<DETALLE_ARQUEO>();
        //DETALLE_VP
        public List<DETALLE_VP> detalle_vp = new List<DETALLE_VP>();
        //DETALLE_REND
        public List<DETALLE_REND> detalle_rend = new List<DETALLE_REND>();
        ConexSAP connectorSap = new ConexSAP();
        public List<ESTATUS> T_Retorno = new List<ESTATUS>();

        public void cierrecajadefinitivo(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER
            , string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_USUARIO, string P_PAIS, string P_ID_APERTURA, string P_MONTO_CIERRE
            , string P_MONTO_DIF, string P_COMENTARIO_DIF, string P_COMENTARIO_CIERRE, string P_TOTAL_APERTURA, string P_IND_CIERRE, string P_ID_ARQUEO_IN)
        {
            try
            {
                ESTATUS retorno;
                //CAB_ARQUEO cabarqueo;
                RESUMEN_VP resumenvp;
                DETALLE_ARQUEO detarqueo;
                DETALLE_REND detallerend;
                DETALLE_VP detallevp;
                DET_EFECTIVO detefectivo;
                T_Retorno.Clear();
                det_arqueo.Clear();
                detalle_rend.Clear();
                det_efectivo.Clear();
                //cab_arqueo.Clear();
                resumen_viapago.Clear();
                detalle_vp.Clear();
                errormessage = "";
                message = "";
                diasatraso = "";
                numerocierre = "";
                
                MontoIngresos = 0;
                MontoEfect = 0;
                MontoChqDia = 0;
                MontoChqFech = 0;
                MontoTransf = 0;
                MontoValeV = 0;
                MontoDepot = 0;
                MontoTarj = 0;
                MontoFinanc = 0;
                MontoApp = 0;
                MontoCredit = 0;
                MontoEgresos = 0;
                MontoCCurse = 0;
                MontoFondosFijos = 0;
                SaldoTotal = 0;

                IRfcTable ls_RETORNO;
                IRfcTable lt_DETALLE_REND;
                IRfcTable lt_RESUMEN_VP;
                IRfcTable lt_DET_EFECTIVO;
                IRfcTable lt_DET_ARQUEO;
                IRfcTable lt_DETALLE_VP;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_FM_SAVE_CIERRE");

                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_USUARIO);
                    BapiGetUser.SetValue("LAND", P_PAIS);
                    if (P_MONTO_CIERRE == "")
                        P_MONTO_CIERRE = "0";
                    BapiGetUser.SetValue("MONTO_CIERRE", P_MONTO_CIERRE); //Total caja
                    if (P_MONTO_DIF == "")
                        P_MONTO_DIF = "0";
                    BapiGetUser.SetValue("MONTO_DIF", P_MONTO_DIF); //Diferencia
                    BapiGetUser.SetValue("COMENTARIO_DIF", P_COMENTARIO_DIF); //text
                    BapiGetUser.SetValue("COMENTARIO_CIERRE", P_COMENTARIO_CIERRE); //text
                    BapiGetUser.SetValue("TOTAL_APERTURA", P_TOTAL_APERTURA); //log apertura
                    BapiGetUser.SetValue("IND_CIERRE", P_IND_CIERRE); //"C"
                    BapiGetUser.SetValue("ID_ARQUEO_IN", P_ID_ARQUEO_IN); //buscar numero de arqueo
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA); //log de apertura



                    BapiGetUser.Invoke(SapRfcDestination);

                    diasatraso = BapiGetUser.GetString("DIAS_ATRASO");
                    numerocierre = BapiGetUser.GetString("NUM_CIERRE");
                   

                    lt_DETALLE_VP = BapiGetUser.GetTable("DETALLE_VP");

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
                            detalle_vp.Add(detallevp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_DETALLE_REND = BapiGetUser.GetTable("DETALLE_REND");
                    try
                    {
                        for (int i = 0; i < lt_DETALLE_REND.Count(); i++)
                        {
                            lt_DETALLE_REND.CurrentIndex = i;
                            detallerend = new DETALLE_REND();
                            detallerend.N_VENTA = lt_DETALLE_REND.GetString("N_VENTA");
                            detallerend.FEC_EMI = lt_DETALLE_REND.GetString("FEC_EMI");
                            detallerend.FEC_VENC = lt_DETALLE_REND.GetString("FEC_VENC");
                            detallerend.MONTO = lt_DETALLE_REND.GetString("MONTO");
                            detallerend.NAME1 = lt_DETALLE_REND.GetString("NAME1");
                            detallerend.MONTO_EFEC = lt_DETALLE_REND.GetString("MONTO_EFEC");
                            MontoEfect = MontoEfect + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_EFEC"));
                            detallerend.NUM_CHEQUE = lt_DETALLE_REND.GetString("NUM_CHEQUE");
                            detallerend.MONTO_DIA = lt_DETALLE_REND.GetString("MONTO_DIA");
                            MontoChqDia = MontoChqDia + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_DIA"));
                            detallerend.MONTO_FECHA = lt_DETALLE_REND.GetString("MONTO_FECHA");
                            MontoChqFech = MontoChqFech + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_FECHA"));
                            detallerend.MONTO_TRANSF = lt_DETALLE_REND.GetString("MONTO_TRANSF");
                            MontoTransf = MontoTransf + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_TRANSF"));
                            detallerend.MONTO_VALE_V = lt_DETALLE_REND.GetString("MONTO_VALE_V");
                            MontoValeV = MontoValeV + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_VALE_V"));
                            detallerend.MONTO_DEP = lt_DETALLE_REND.GetString("MONTO_DEP");
                            MontoDepot = MontoDepot + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_DEP"));
                            detallerend.MONTO_TARJ = lt_DETALLE_REND.GetString("MONTO_TARJ");
                            MontoTarj = MontoTarj + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_TARJ"));
                            detallerend.MONTO_FINANC = lt_DETALLE_REND.GetString("MONTO_FINANC");
                            MontoFinanc = MontoFinanc + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_FINANC"));
                            detallerend.MONTO_APP = lt_DETALLE_REND.GetString("MONTO_APP");
                            MontoApp = MontoApp + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_APP"));
                            detallerend.MONTO_CREDITO = lt_DETALLE_REND.GetString("MONTO_CREDITO");
                            MontoCredit = MontoCredit + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_CREDITO"));
                            detallerend.PATENTE = lt_DETALLE_REND.GetString("PATENTE");
                            detallerend.MONTO_C_CURSE = lt_DETALLE_REND.GetString("MONTO_C_CURSE");
                            MontoCCurse = MontoCCurse + Convert.ToDouble(lt_DETALLE_REND.GetString("MONTO_C_CURSE"));
                            detallerend.DOC_SAP = lt_DETALLE_REND.GetString("DOC_SAP");
                            detalle_rend.Add(detallerend);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    MontoIngresos = MontoIngresos + MontoEfect + MontoChqDia + MontoChqFech + MontoTransf + MontoValeV +
                        MontoDepot + MontoTarj + MontoFinanc + MontoApp + MontoCredit + MontoCCurse;
                    SaldoTotal = MontoIngresos - MontoEgresos;


                    lt_RESUMEN_VP = BapiGetUser.GetTable("RESUMEN_VP");
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

                    //lt_CAB_ARQUEO = BapiGetUser.GetTable("DET_EFECTIVO");
                    //try
                    //{
                    //    for (int i = 0; i < lt_CAB_ARQUEO.Count(); i++)
                    //    {
                    //        lt_CAB_ARQUEO.CurrentIndex = i;
                    //        cabarqueo = new CAB_ARQUEO();
                    //        cabarqueo.MANDT = lt_CAB_ARQUEO.GetString("MANDT");
                    //        cabarqueo.LAND = lt_CAB_ARQUEO.GetString("LAND");
                    //        cabarqueo.ID_ARQUEO = lt_CAB_ARQUEO.GetString("ID_ARQUEO");
                    //        cabarqueo.ID_REGISTRO = lt_CAB_ARQUEO.GetString("ID_DENOMINACION");
                    //        cabarqueo.ID_CAJA = lt_CAB_ARQUEO.GetString("CANTIDAD");
                    //        cabarqueo.MONTO_CIERRE = lt_CAB_ARQUEO.GetString("MONTO_TOTAL");
                    //        cabarqueo.MONTO_DIF = lt_CAB_ARQUEO.GetString("ID_DENOMINACION");
                    //        cabarqueo.COMENTARIO_DIF = lt_CAB_ARQUEO.GetString("CANTIDAD");
                    //        cabarqueo.NULO = lt_CAB_ARQUEO.GetString("MONTO_TOTAL");
                    //        cab_arqueo.Add(cabarqueo);
                    //    }
                    //}
                    //catch (Exception ex)
                    //{
                    //    Console.WriteLine(ex.Message, ex.StackTrace);
                    //    MessageBox.Show(ex.Message + ex.StackTrace);
                    //}

                    //lt_DET_ARQUEO = BapiGetUser.GetTable("DET_ARQUEO");
                    //try
                    //{
                    //    for (int i = 0; i < lt_DET_ARQUEO.Count(); i++)
                    //    {
                    //        lt_DET_ARQUEO.CurrentIndex = i;
                    //        detarqueo = new DET_ARQUEO();
                    //        detarqueo.MANDT = lt_DET_ARQUEO.GetString("MANDT");
                    //        detarqueo.LAND = lt_DET_ARQUEO.GetString("LAND");
                    //        detarqueo.ID_ARQUEO = lt_DET_ARQUEO.GetString("ID_ARQUEO");
                    //        detarqueo.ID_DENOMINACION = lt_DET_ARQUEO.GetString("ID_DENOMINACION");
                    //        detarqueo.CANTIDAD = lt_DET_ARQUEO.GetString("CANTIDAD");
                    //        detarqueo.MONTO_TOTAL = lt_DET_ARQUEO.GetString("MONTO_TOTAL");
                    //        det_arqueo.Add(detarqueo);
                    //    }
                    //}
                    //catch (Exception ex)
                    //{
                    //    Console.WriteLine(ex.Message, ex.StackTrace);
                    //    MessageBox.Show(ex.Message + ex.StackTrace);
                    //}

                    lt_DET_ARQUEO = BapiGetUser.GetTable("DETALLE_ARQUEO");
                    try
                    {
                        for (int i = 0; i < lt_DETALLE_REND.Count(); i++)
                        {
                            lt_DET_ARQUEO.CurrentIndex = i;
                            detarqueo = new DETALLE_ARQUEO();
                            detarqueo.LAND = lt_DET_ARQUEO.GetString("LAND");
                            detarqueo.ID_CAJA = lt_DET_ARQUEO.GetString("ID_CAJA");
                            detarqueo.USUARIO = lt_DET_ARQUEO.GetString("USUARIO");
                            detarqueo.SOCIEDAD = lt_DET_ARQUEO.GetString("SOCIEDAD");
                            detarqueo.HORA_REND = lt_DET_ARQUEO.GetString("HORA_REND");
                            detarqueo.FECHA_REND = lt_DET_ARQUEO.GetString("FECHA_REND");
                            detarqueo.MONEDA = lt_DET_ARQUEO.GetString("MONEDA");
                            detarqueo.VIA_PAGO = lt_DET_ARQUEO.GetString("VIA_PAGO");
                            detarqueo.TIPO_MONEDA = lt_DET_ARQUEO.GetString("TIPO_MONEDA");
                            detarqueo.CANTIDAD_MON = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                            detarqueo.SUMA_MON_BILL = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                            detarqueo.CANTIDAD_DOC = lt_DET_ARQUEO.GetString("CANTIDAD_DOC");
                            detarqueo.SUMA_DOCS = lt_DET_ARQUEO.GetString("MONTO_DSUMA_DOCSEP");
                            det_arqueo.Add(detarqueo);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_DET_EFECTIVO = BapiGetUser.GetTable("DET_EFECTIVO");
                    try
                    {
                        for (int i = 0; i < lt_DET_EFECTIVO.Count(); i++)
                        {
                            lt_DET_ARQUEO.CurrentIndex = i;
                            detefectivo = new DET_EFECTIVO();
                            detefectivo.MANDT = lt_DET_EFECTIVO.GetString("MANDT");
                            detefectivo.LAND = lt_DET_EFECTIVO.GetString("LAND");
                            detefectivo.ID_ARQUEO = lt_DET_EFECTIVO.GetString("ID_ARQUEO");
                            detefectivo.ID_DENOMINACION = lt_DET_EFECTIVO.GetString("ID_DENOMINACION");
                            detefectivo.CANTIDAD = lt_DET_EFECTIVO.GetString("CANTIDAD");
                            detefectivo.MONTO_TOTAL = lt_DET_EFECTIVO.GetString("MONTO_TOTAL");
                            det_efectivo.Add(detefectivo);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }


                    ls_RETORNO = BapiGetUser.GetTable("ESTATUS");
                    try
                    {
                        for (int i = 0; i < ls_RETORNO.Count(); i++)
                        {
                            ls_RETORNO.CurrentIndex = i;
                            retorno = new ESTATUS();
                            if (ls_RETORNO.GetString("TYPE") == "S")
                            {
                                message = message + " - " + ls_RETORNO.GetString("MESSAGE");
                                //numerocierre = ls_RETORNO.GetString("MESSAGE_V4");
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
