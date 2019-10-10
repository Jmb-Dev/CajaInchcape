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
using CajaIndigo.AppPersistencia.Class.ArqueoCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;


namespace CajaIndigo.AppPersistencia.Class.RendicionCaja
{
    class RendicionCaja
    {
        //public List<T_DOCUMENTOS> ObjDatosMonitor = new List<T_DOCUMENTOS>();
        public string errormessage = "";
        public string message = "";
        public string IdCaja = "";
        public string id_arqueo = "";
        public string id_cierre = "";
        public string FechArqueo = "";
        public string FechaArqueoHasta = "";
        public string CajeroResp = "";
        public string Sucursal = "";
        public string NombreCaja = "";
        public string Sociedad = "";
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
        public double MontoCCurse = 0;
        public double MontoCredit = 0;
        public double MontoEgresos = 0;
        public double MontoFondosFijos = 0;
        public double SaldoTotal = 0;
        public string M1 = "0";
        public string M5 = "0";
        public string M10 = "0";
        public string M50 = "0";
        public string M100 = "0";
        public string M500 = "0";
        public string M1000 = "0";
        public string M2000 = "0";
        public string M5000 = "0";
        public string M10000 = "0";
        public string M20000 = "0";
        public string C1 = "0";
        public string C5 = "0";
        public string C10 = "0";
        public string C50 = "0";
        public string C100 = "0";
        public string C500 = "0";
        public string C1000 = "0";
        public string C2000 = "0";
        public string C5000 = "0";
        public string C10000 = "0";
        public string C20000 = "0";
        //DETALLE_VP
        public List<DETALLE_VP> detalle_viapago = new List<DETALLE_VP>(); 
        //RESUMEN_VP
        public List<RESUMEN_VP> resumen_viapago = new List<RESUMEN_VP>();
        //DETALLE_REND
        public List<DETALLE_REND> detalle_rend = new List<DETALLE_REND>();
        //DETALLE_ARQUEO
        public List<DETALLE_ARQUEO> detalle_arqueo = new List<DETALLE_ARQUEO>();
        ConexSAP connectorSap = new ConexSAP();
        public List<ESTATUS> T_Retorno = new List<ESTATUS>();

        public List<INFO_SOCI> ObjInfoSoc = new List<INFO_SOCI>();

        public void rendicioncaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_DATUMDESDE, string P_DATUMHASTA, string P_USUARIO
            , string P_PAIS, string P_SOCIEDAD, string P_ID_APERTURA,  string P_ID_CIERRE, string P_IND_ARQUEO, string P_MONEDA)
        {
            try
            {
                ESTATUS retorno;
                DETALLE_VP detallevp;
                RESUMEN_VP resumenvp;
                DETALLE_REND detallerend;
                DETALLE_ARQUEO detallearqueo;
                INFO_SOCI InfoSociedad;
              
                T_Retorno.Clear();
                detalle_rend.Clear();
                detalle_arqueo.Clear();
                detalle_viapago.Clear();
                resumen_viapago.Clear();
                errormessage = "";
                message = "";
                IdCaja = "";
                FechArqueo = "";
                FechaArqueoHasta = "";
                CajeroResp = "";
                Sucursal = "";
                Sociedad = "";
                NombreCaja = "";
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
                MontoFondosFijos = 0;
                SaldoTotal = 0;
                IRfcStructure ls_RETORNO;
                IRfcTable lt_DETALLE_VP;
                IRfcTable lt_RESUMEN_VP;
                IRfcTable lt_DETALLE_REND;
                IRfcTable lt_DET_ARQUEO;
                IRfcTable lt_INFO_SOCI;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_REP_REN_CAJA");

                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_USUARIO);
                    BapiGetUser.SetValue("ID_CIERRE", P_ID_CIERRE);
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);
                    BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                    BapiGetUser.SetValue("IND_ARQUEO", P_IND_ARQUEO);
                    BapiGetUser.SetValue("LAND", P_PAIS);

                    BapiGetUser.Invoke(SapRfcDestination);

                    BapiGetUser.GetString("ID_CAJA_OUT");
                    FechArqueo = BapiGetUser.GetString("DATE_ARQ_OUT");
                    FechaArqueoHasta = BapiGetUser.GetString("DATE_ARQ_TO_OUT");
                    CajeroResp = BapiGetUser.GetString("CAJERO_RESP_OUT");
                    Sucursal = BapiGetUser.GetString("SUCURSAL");
                    NombreCaja = BapiGetUser.GetString("NOM_CAJA");
                    Sociedad = BapiGetUser.GetString("SOCIEDAD_OUT");
                    id_arqueo = BapiGetUser.GetString("ID_ARQUEO");
                    id_cierre = BapiGetUser.GetString("ID_CIERRE_OUT");
                    lt_DETALLE_VP = BapiGetUser.GetTable("DETALLE_VP");
                    lt_INFO_SOCI = BapiGetUser.GetTable("INFO_SOCI");
                    try
                    {
                        for (int i = 0; i < lt_INFO_SOCI.Count(); i++)
                        {
                            lt_INFO_SOCI.CurrentIndex = i;
                            InfoSociedad = new INFO_SOCI();
                            InfoSociedad.BUKRS = lt_INFO_SOCI.GetString("BUKRS");
                            InfoSociedad.BUTXT = lt_INFO_SOCI.GetString("BUTXT");
                            InfoSociedad.STCD1 = lt_INFO_SOCI.GetString("STCD1");
                            ObjInfoSoc.Add(InfoSociedad);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }
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
                            FormatoMonedas FM = new FormatoMonedas();
                            if (lt_DETALLE_VP[i].GetString("MONEDA") == "CLP")
                            {
                                string Formateo = FM.FormatoMonedaChilena(lt_DETALLE_VP[i].GetString("MONTO").Trim(), "2");
                                detallevp.MONTO = Formateo;
                            }
                            else
                            {
                                string Formateo = FM.FormatoMonedaExtranjera(lt_DETALLE_VP[i].GetString("MONTO").Trim());
                                detallevp.MONTO = Formateo;
                            }
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
                            detalle_viapago.Add(detallevp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

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

                            //string Valor = lt_RESUMEN_VP[i].GetString("MONTO_EFEC").Trim();
                            ////if (lt_RESUMEN_VP.GetString("VIA_PAGO") == "0" || lt_RESUMEN_VP.GetString("VIA_PAGO") == "N")
                            ////{
                            ////    MontoEgresos = MontoEgresos + Convert.ToDouble(lt_RESUMEN_VP.GetString("MONTO"));
                            ////}
                            resumenvp.TEXT1 = lt_RESUMEN_VP.GetString("TEXT1");
                            resumenvp.MONEDA = lt_RESUMEN_VP.GetString("MONEDA");
                            FormatoMonedas FM = new FormatoMonedas();
                            if (lt_RESUMEN_VP[i].GetString("MONEDA") == "CLP")
                            {
                                resumenvp.MONTO = FM.FormatoMonedaChilena(lt_RESUMEN_VP[i].GetString("MONTO"),  "1");
                            }
                            else
                            {
                                resumenvp.MONTO = FM.FormatoMonedaExtranjera(lt_RESUMEN_VP[i].GetString("MONTO"));
                            }
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
                            detallerend.MONEDA = lt_DETALLE_REND.GetString("MONEDA");
                            detallerend.VIA_PAGO = lt_DETALLE_REND.GetString("VIA_PAGO");

                            FormatoMonedas FM = new FormatoMonedas();
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                               detallerend.MONTO = FM.FormatoMonedaChilena(lt_DETALLE_REND[i].GetString("MONTO"),  "2");
                            }
                            else
                            {
                               detallerend.MONTO = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO"));
                            }
                            detallerend.NAME1 = lt_DETALLE_REND.GetString("NAME1");
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                               MontoEfect = MontoEfect + Convert.ToDouble(FM.FormatoMonedaChilena(lt_DETALLE_REND[i].GetString("MONTO_EFEC"),  "2"));
                               detallerend.MONTO_EFEC = FM.FormatoMonedaChilena(lt_DETALLE_REND[i].GetString("MONTO_EFEC"), "2");
                               string Valor = lt_DETALLE_REND[i].GetString("MONTO_EFEC").Trim();

                               if (Valor.Contains("-"))
                               {
                                   MontoEgresos = MontoEgresos + Convert.ToDouble(lt_DETALLE_REND[i].GetString("MONTO_EFEC"));
                               }                         
                            }
                            else
                            {
                                detallerend.MONTO_EFEC = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_EFEC"));
                            }
                            detallerend.NUM_CHEQUE = lt_DETALLE_REND.GetString("NUM_CHEQUE");
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_DIA = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_DIA"),  "2");
                                MontoChqDia = MontoChqDia + Convert.ToDouble(FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_DIA"),  "2"));
                            }
                            else
                            {
                                detallerend.MONTO_DIA = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_DIA"));
                                MontoChqDia = MontoChqDia + Convert.ToDouble(detallerend.MONTO_DIA);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                MontoChqFech = MontoChqFech + Convert.ToDouble(FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_FECHA"), "2"));
                                detallerend.MONTO_FECHA = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_FECHA"), "2");
                            }
                            else
                            {
                                detallerend.MONTO_FECHA = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_FECHA"));
                                MontoChqFech = MontoChqFech + Convert.ToDouble(detallerend.MONTO_FECHA);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_TRANSF = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_TRANSF"), "2");
                                MontoTransf = MontoTransf + Convert.ToDouble(FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_TRANSF"), "2"));
                            }
                            else
                            {
                                detallerend.MONTO_TRANSF = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_TRANSF"));
                                MontoTransf = MontoTransf + Convert.ToDouble(detallerend.MONTO_TRANSF);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_VALE_V = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_VALE_V"), "2");
                                MontoValeV = MontoValeV + Convert.ToDouble(detallerend.MONTO_VALE_V);
                            }
                            else
                            {
                                detallerend.MONTO_VALE_V = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_VALE_V"));
                                MontoValeV = MontoValeV + Convert.ToDouble(detallerend.MONTO_VALE_V);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_DEP = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_DEP"), "2");
                                MontoDepot = MontoDepot + Convert.ToDouble(detallerend.MONTO_DEP);
                            }
                            else
                            {
                                detallerend.MONTO_DEP = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_DEP"));
                                MontoDepot = MontoDepot + Convert.ToDouble(detallerend.MONTO_DEP);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_TARJ = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_TARJ"), "2");
                                MontoTarj = MontoTarj + Convert.ToDouble(detallerend.MONTO_TARJ);
                            }
                            else
                            {
                                detallerend.MONTO_TARJ = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_TARJ"));
                                MontoTarj = MontoTarj + Convert.ToDouble(detallerend.MONTO_TARJ);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_FINANC = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_FINANC"), "2");
                                MontoFinanc = MontoFinanc + Convert.ToDouble(detallerend.MONTO_FINANC);
                            }
                            else
                            {
                                detallerend.MONTO_FINANC = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_FINANC"));
                                MontoFinanc = MontoFinanc + Convert.ToDouble(detallerend.MONTO_FINANC);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_APP = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_APP"), "2");
                                MontoApp = MontoApp + Convert.ToDouble(detallerend.MONTO_APP);
                            }
                            else
                            {
                                detallerend.MONTO_APP = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_APP"));
                                MontoApp = MontoApp + Convert.ToDouble(detallerend.MONTO_APP);
                            }
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_CREDITO = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_CREDITO"), "2");
                                MontoCredit = MontoCredit + Convert.ToDouble(detallerend.MONTO_CREDITO);
                            }
                            else
                            {
                                detallerend.MONTO_CREDITO = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_CREDITO"));
                                MontoCredit = MontoCredit + Convert.ToDouble(detallerend.MONTO_CREDITO);
                            }
                            detallerend.PATENTE = lt_DETALLE_REND.GetString("PATENTE");
                            if (detallerend.MONEDA == P_MONEDA)
                            {
                                detallerend.MONTO_C_CURSE = FM.FormatoMonedaChilena(lt_DETALLE_REND.GetString("MONTO_C_CURSE"), "2");
                                MontoCCurse = MontoCCurse + Convert.ToDouble(detallerend.MONTO_C_CURSE);
                            }
                            else
                            {
                                detallerend.MONTO_C_CURSE = FM.FormatoMonedaExtranjera(lt_DETALLE_REND[i].GetString("MONTO_C_CURSE"));
                                MontoCCurse = MontoCCurse + Convert.ToDouble(detallerend.MONTO_C_CURSE);
                            }
                            detallerend.DOC_SAP = lt_DETALLE_REND.GetString("DOC_SAP");
                            detalle_rend.Add(detallerend);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    MontoIngresos =  MontoIngresos + MontoEfect + MontoChqDia + MontoChqFech + MontoTransf + MontoValeV + 
                        MontoDepot + MontoTarj + MontoFinanc + MontoApp + MontoCredit + MontoCCurse;
                    SaldoTotal = MontoIngresos; // +MontoEgresos;

                    lt_DET_ARQUEO = BapiGetUser.GetTable("DETALLE_ARQUEO");
                    try
                    {
                        for (int i = 0; i < lt_DET_ARQUEO.Count(); i++)
                        {
                            lt_DET_ARQUEO.CurrentIndex = i;
                            detallearqueo = new DETALLE_ARQUEO();
                            detallearqueo.LAND = lt_DET_ARQUEO.GetString("LAND");
                            detallearqueo.ID_CAJA = lt_DET_ARQUEO.GetString("ID_CAJA");
                            detallearqueo.USUARIO = lt_DET_ARQUEO.GetString("USUARIO");
                            detallearqueo.SOCIEDAD = lt_DET_ARQUEO.GetString("SOCIEDAD");
                            detallearqueo.FECHA_REND = lt_DET_ARQUEO.GetString("FECHA_REND");
                            detallearqueo.HORA_REND = lt_DET_ARQUEO.GetString("HORA_REND");
                            detallearqueo.MONEDA = lt_DET_ARQUEO.GetString("MONEDA");
                            detallearqueo.VIA_PAGO = lt_DET_ARQUEO.GetString("VIA_PAGO");
                            detallearqueo.TIPO_MONEDA = lt_DET_ARQUEO.GetString("TIPO_MONEDA");
                            detallearqueo.CANTIDAD_MON = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                            //*
                            if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                            {
                                string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                if (Valor.Contains("-"))
                                {
                                    Valor = "-" + Valor.Replace("-", "");
                                }
                                Valor = Valor.Replace(".", "");
                                Valor = Valor.Replace(",", "");
                                decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                detallearqueo.SUMA_MON_BILL = Cualquiernombre;
                            }
                            else
                            {
                                string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                decimal ValorAux = Convert.ToDecimal(moneda);
                                detallearqueo.SUMA_MON_BILL = string.Format("{0:0,0.##}", ValorAux);
                            }
                            if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                            {
                                string Valor = lt_DET_ARQUEO[i].GetString("SUMA_DOCS").Trim();
                                if (Valor.Contains("-"))
                                {
                                    Valor = "-" + Valor.Replace("-", "");
                                }
                                Valor = Valor.Replace(".", "");
                                Valor = Valor.Replace(",", "");
                                decimal ValorAux = Convert.ToDecimal(Valor.Substring(0,Valor.Length-2));
                                string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                detallearqueo.SUMA_DOCS = Cualquiernombre;
                            }
                            else
                            {
                                string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_DOCS"));
                                decimal ValorAux = Convert.ToDecimal(moneda);
                                detallearqueo.SUMA_DOCS = string.Format("{0:0,0.##}", ValorAux);
                            }
                            detallearqueo.SUMA_DOCS = lt_DET_ARQUEO.GetString("SUMA_DOCS");
                            C1 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000000001")
                            {
                                //C1 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M1 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M1 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //M1 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                             }
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000000005")
                            {
                                //C5 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M5 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M5 = string.Format("{0:0,0.##}", ValorAux);
                                }
                             }
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000000010")
                             {
                                //C10 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M10 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M10 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //M10 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                             }
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000000050")
                            {
                                //C50 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M50 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M50= string.Format("{0:0,0.##}", ValorAux);
                                }
                                //M50 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                             }
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000000100")
                            {
                                //C100 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M100 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M100 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //M100 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                             }
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000000500")
                             {
                                 //C500 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                 if (lt_DET_ARQUEO[i].GetString("MONEDA") == "CLP")
                                 {
                                     string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                     if (Valor.Contains("-"))
                                     {
                                         Valor = "-" + Valor.Replace("-", "");
                                     }
                                     Valor = Valor.Replace(".", "");
                                     Valor = Valor.Replace(",", "");
                                     decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                     string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                     M500 = Cualquiernombre;
                                 }
                                 else
                                 {
                                     string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                     decimal ValorAux = Convert.ToDecimal(moneda);
                                     M500 = string.Format("{0:0,0.##}", ValorAux);
                                 }
                                 //M500 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                             }
                             if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000001000")
                             {
                                 //C1000 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                 if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                 {
                                     string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                     if (Valor.Contains("-"))
                                     {
                                         Valor = "-" + Valor.Replace("-", "");
                                     }
                                     Valor = Valor.Replace(".", "");
                                     Valor = Valor.Replace(",", "");
                                     decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                     string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                     M1000 = Cualquiernombre;
                                 }
                                 else
                                 {
                                     string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                     decimal ValorAux = Convert.ToDecimal(moneda);
                                     M1000 = string.Format("{0:0,0.##}", ValorAux);
                                 }
                                 //M1000 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                             }
                            if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000002000")
                            {
                                //C2000 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M2000 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M2000 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                //M2000 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                            }
                            if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000005000")
                            {
                                //C5000 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M5000 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M5000 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                M5000 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                            }
                            if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000010000")
                            {
                                //C10000 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M10000 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M10000 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                M10000 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                            }
                            if (lt_DET_ARQUEO.GetString("TIPO_MONEDA") == "0000020000")
                            {
                                //C20000 = lt_DET_ARQUEO.GetString("CANTIDAD_MON");
                                if (lt_DET_ARQUEO[i].GetString("MONEDA") == P_MONEDA)
                                {
                                    string Valor = lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    M20000 = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_DET_ARQUEO[i].GetString("SUMA_MON_BILL"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    M20000 = string.Format("{0:0,0.##}", ValorAux);
                                }
                                M20000 = lt_DET_ARQUEO.GetString("SUMA_MON_BILL");
                            }
                            detalle_arqueo.Add(detallearqueo);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    ls_RETORNO = BapiGetUser.GetStructure("RETORNO");
                    try
                    {
                        for (int i = 0; i < ls_RETORNO.Count(); i++)
                        {
                            //ls_RETORNO.CurrentIndex = i;
                            retorno = new ESTATUS();
                            if (ls_RETORNO.GetString("TYPE") == "S")
                            {
                                message = message + " - " + ls_RETORNO.GetString("MESSAGE");
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
