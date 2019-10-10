using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes.Estructura;
using CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.ReportesCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.PreCierreCaja.Estructura;

namespace CajaIndigo.AppPersistencia.Class.ReportesCaja
{
    class ReportesCaja
    {
          
       //public List<T_DOCUMENTOS> ObjDatosMonitor = new List<T_DOCUMENTOS>();
        public string errormessage = "";
        public string message = "";
        public string IdCaja = "";
        public string Pais = "";
        public string Caja = "";
        public string RUT = "";
        public string FechArqueo = "";
        public string Tipo = "";
        public string FechaArqueoHasta = "";
        public string CajeroResp = "";
        public string Sucursal = "";
        public string NombreCaja = "";
        public string Sociedad = "";
        public string SociedadR = "";
        public string Empresa = "";
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
        //RESUMEN_VP
        public List<CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura.RESUMEN_VP> resumen_viapago = new List<CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura.RESUMEN_VP>();
        //DETALLE_REND
        public List<DETALLE_REND> detalle_rend = new List<DETALLE_REND>();
        //CABECERA_ARQUEO
        public List<CAB_ARQUEO> cab_arqueo = new List<CAB_ARQUEO>();
        //DETALLE_ARQUEO
        public List<DET_ARQUEO> detalle_arqueo = new List<DET_ARQUEO>();
        //RESUMEN_CAJA
        public List<RESUMEN_CAJA> resumen_caja = new List<RESUMEN_CAJA>();
        //RENDICION_CAJA
        public List<RENDICION_CAJA> rendicion_caja = new List<RENDICION_CAJA>();
        //RESUMEN_MENSUAL
        public List<RESUMEN_MENSUAL> resumen_mensual = new List<RESUMEN_MENSUAL>();
        //INFO_SOC
        public List<INFO_SOC> info_soc = new List<INFO_SOC>();
        ConexSAP connectorSap = new ConexSAP();
        public List<RETORNO> T_Retorno = new List<RETORNO>();

        public void reportescaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER
            , string P_IDIOMA, string P_ID_CAJA, string P_DATUMDESDE, string P_DATUMHASTA, string P_USUARIO
            , string P_PAIS, string P_SOCIEDAD, string P_ID_APERTURA, string P_ID_CIERRE, string P_ID_REPORT)
        {
            try
            {
                T_Retorno.Clear();
                resumen_viapago.Clear();
                detalle_rend.Clear();
                cab_arqueo.Clear();
                detalle_arqueo.Clear();
                resumen_caja.Clear();
                rendicion_caja.Clear();
                resumen_mensual.Clear();
                info_soc.Clear();
                RETORNO retorno;
                CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura.RESUMEN_VP resumenvp;
                DETALLE_REND detallerend;
                CAB_ARQUEO cabarqueo;
                DET_ARQUEO detallearqueo;
                RESUMEN_CAJA resumencaja;
                RESUMEN_MENSUAL resumenmensual;
                RENDICION_CAJA rendicioncaja;
                INFO_SOC infosoc;
              
                errormessage = "";
                message = "";
                IdCaja = "";
                FechArqueo = "";
                FechaArqueoHasta = "";
                CajeroResp = "";
                RUT = "";
                Sucursal = "";
                Sociedad = "";
                SociedadR = "";
                Empresa = "";
                NombreCaja = "";
                Tipo = "";
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
                IRfcTable lt_RESUMEN_VP;
                IRfcTable lt_DETALLE_REND;
                IRfcTable lt_CAB_ARQUEO;
                IRfcTable lt_DET_ARQUEO;
                IRfcTable lt_RESUMEN_CAJA;
                IRfcTable lt_RENDICION_CAJA;
                IRfcTable lt_RESUMEN_MENSUAL;
                IRfcTable lt_INFOSOC;
                
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_REP_CAJA");

                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                   
                        BapiGetUser.SetValue("DATE_DOC_FROM", Convert.ToDateTime(P_DATUMDESDE));
                        BapiGetUser.SetValue("DATE_DOC_TO", Convert.ToDateTime(P_DATUMHASTA));
                   
                    BapiGetUser.SetValue("USUARIO", P_USUARIO);
                    BapiGetUser.SetValue("ID_CIERRE", P_ID_CIERRE);
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);
                    BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                    BapiGetUser.SetValue("ID_REPORT", P_ID_REPORT);
                    BapiGetUser.SetValue("LAND", P_PAIS);
                    Tipo = P_ID_REPORT;
                    FechArqueo = P_DATUMDESDE;
                    FechaArqueoHasta = P_DATUMHASTA;

                    BapiGetUser.Invoke(SapRfcDestination);

                    Caja = BapiGetUser.GetString("ID_CAJA_OUT");
                    CajeroResp = BapiGetUser.GetString("CAJERO_RESP_OUT");
                    Sucursal = BapiGetUser.GetString("SUCURSAL");
                    NombreCaja = BapiGetUser.GetString("NOM_CAJA");
                    Sociedad = BapiGetUser.GetString("SOCIEDAD_OUT");
                    Pais = BapiGetUser.GetString("LAND_OUT");


                    lt_INFOSOC = BapiGetUser.GetTable("INFO_SOC");
                    try
                    {
                        for (int i = 0; i < lt_INFOSOC.Count(); i++)
                        {
                            lt_INFOSOC.CurrentIndex = i;
                            infosoc = new INFO_SOC();
                            SociedadR = lt_INFOSOC.GetString("BUKRS");
                            infosoc.BUKRS = lt_INFOSOC.GetString("BUKRS");
                            Empresa = lt_INFOSOC.GetString("BUTXT");
                            infosoc.BUTXT = lt_INFOSOC.GetString("BUTXT");
                            RUT = lt_INFOSOC.GetString("STCD1");
                            infosoc.STCD1 = lt_INFOSOC.GetString("STCD1");
                           

                        info_soc.Add(infosoc);
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
                            resumenvp = new CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura.RESUMEN_VP();
                            resumenvp.LAND = lt_RESUMEN_VP.GetString("LAND");
                            resumenvp.ID_CAJA = lt_RESUMEN_VP.GetString("ID_CAJA");
                            SociedadR = lt_RESUMEN_VP.GetString("SOCIEDAD");
                            resumenvp.SOCIEDAD = lt_RESUMEN_VP.GetString("SOCIEDAD");
                            Empresa = lt_RESUMEN_VP.GetString("SOCIEDAD_TXT");
                            resumenvp.SOCIEDAD_TXT = lt_RESUMEN_VP.GetString("SOCIEDAD_TXT");
                            resumenvp.VIA_PAGO = lt_RESUMEN_VP.GetString("VIA_PAGO");
                            if (lt_RESUMEN_VP.GetString("VIA_PAGO") == "N"|| lt_RESUMEN_VP.GetString("VIA_PAGO") == "0")
                            {
                                MontoEgresos = MontoEgresos + Convert.ToDouble(lt_RESUMEN_VP.GetString("MONTO"));
                            }
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

                    MontoIngresos =  MontoIngresos + MontoEfect + MontoChqDia + MontoChqFech + MontoTransf + MontoValeV + 
                        MontoDepot + MontoTarj + MontoFinanc + MontoApp + MontoCredit + MontoCCurse;
                    SaldoTotal = MontoIngresos; // +MontoEgresos;

                    lt_CAB_ARQUEO = BapiGetUser.GetTable("CAB_ARQUEO");
                    try
                    {
                        for (int i = 0; i < lt_CAB_ARQUEO.Count(); i++)
                        {
                            lt_CAB_ARQUEO.CurrentIndex = i;
                            cabarqueo = new CAB_ARQUEO();
                            cabarqueo.MANDT = lt_CAB_ARQUEO.GetString("MANDT");
                            cabarqueo.LAND = lt_CAB_ARQUEO.GetString("LAND");
                            cabarqueo.ID_ARQUEO = lt_CAB_ARQUEO.GetString("ID_ARQUEO");
                            cabarqueo.ID_REGISTRO = lt_CAB_ARQUEO.GetString("ID_DENOMINACION");
                            cabarqueo.ID_CAJA = lt_CAB_ARQUEO.GetString("CANTIDAD");
                            cabarqueo.MONTO_CIERRE = lt_CAB_ARQUEO.GetString("MONTO_TOTAL");
                            cabarqueo.MONTO_DIF = lt_CAB_ARQUEO.GetString("ID_DENOMINACION");
                            cabarqueo.COMENTARIO_DIF = lt_CAB_ARQUEO.GetString("CANTIDAD");
                            cabarqueo.NULO = lt_CAB_ARQUEO.GetString("MONTO_TOTAL");
                            cab_arqueo.Add(cabarqueo);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_DET_ARQUEO = BapiGetUser.GetTable("DET_ARQUEO");
                    try
                    {
                        for (int i = 0; i < lt_DET_ARQUEO.Count(); i++)
                        {
                            lt_DET_ARQUEO.CurrentIndex = i;
                            detallearqueo = new DET_ARQUEO();
                            detallearqueo.MANDT = lt_DET_ARQUEO.GetString("MANDT");
                            detallearqueo.LAND = lt_DET_ARQUEO.GetString("LAND");
                            detallearqueo.ID_ARQUEO = lt_DET_ARQUEO.GetString("ID_ARQUEO");
                            detallearqueo.ID_DENOMINACION = lt_DET_ARQUEO.GetString("ID_DENOMINACION");
                            detallearqueo.CANTIDAD = lt_DET_ARQUEO.GetString("CANTIDAD");
                            detallearqueo.MONTO_TOTAL = lt_DET_ARQUEO.GetString("MONTO_TOTAL");
                            detalle_arqueo.Add(detallearqueo);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_RESUMEN_CAJA = BapiGetUser.GetTable("RESUMEN_CAJA");
                    try
                    {
                        for (int i = 0; i < lt_RESUMEN_CAJA.Count(); i++)
                        {
                            lt_RESUMEN_CAJA.CurrentIndex = i;
                            resumencaja = new RESUMEN_CAJA();
                            resumencaja.ID_SUCURSAL = lt_RESUMEN_CAJA.GetString("ID_SUCURSAL");
                            resumencaja.SUCURSAL = lt_RESUMEN_CAJA.GetString("SUCURSAL");
                            resumencaja.ID_CAJA = lt_RESUMEN_CAJA.GetString("ID_CAJA");
                            resumencaja.NOM_CAJA = lt_RESUMEN_CAJA.GetString("NOM_CAJA");
                            resumencaja.MONTO_EFEC = lt_RESUMEN_CAJA.GetString("MONTO_EFEC");
                            MontoEfect = MontoEfect + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_EFEC"));
                            resumencaja.MONTO_DIA = lt_RESUMEN_CAJA.GetString("MONTO_DIA");
                            MontoChqDia = MontoChqDia + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_DIA"));
                            resumencaja.MONTO_FECHA = lt_RESUMEN_CAJA.GetString("MONTO_FECHA");
                            MontoChqFech = MontoChqFech + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_FECHA"));
                            resumencaja.MONTO_TRANSF = lt_RESUMEN_CAJA.GetString("MONTO_TRANSF");
                            MontoTransf = MontoTransf + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_TRANSF"));
                            resumencaja.MONTO_VALE_V = lt_RESUMEN_CAJA.GetString("MONTO_VALE_V");
                            MontoValeV = MontoValeV + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_VALE_V"));
                            resumencaja.MONTO_DEP = lt_RESUMEN_CAJA.GetString("MONTO_DEP");
                            MontoDepot = MontoDepot + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_DEP"));
                            resumencaja.MONTO_TARJ = lt_RESUMEN_CAJA.GetString("MONTO_TARJ");
                            MontoTarj = MontoTarj + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_TARJ"));
                            resumencaja.MONTO_FINANC = lt_RESUMEN_CAJA.GetString("MONTO_FINANC");
                            MontoFinanc = MontoFinanc + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_FINANC"));
                            resumencaja.MONTO_APP = lt_RESUMEN_CAJA.GetString("MONTO_APP");
                            MontoApp = MontoApp + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_APP"));
                            resumencaja.MONTO_CREDITO = lt_RESUMEN_CAJA.GetString("MONTO_CREDITO");
                            MontoCredit = MontoCredit + Convert.ToDouble(lt_RESUMEN_CAJA.GetString("MONTO_CREDITO"));
                            //resumencaja.MONEDA = lt_RESUMEN_CAJA.GetString("MONEDA");
                            resumen_caja.Add(resumencaja);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_RENDICION_CAJA = BapiGetUser.GetTable("RENDICION_CAJA");
                    try
                    {
                        for (int i = 0; i < lt_RENDICION_CAJA.Count(); i++)
                        {
                            lt_RENDICION_CAJA.CurrentIndex = i;
                            rendicioncaja = new RENDICION_CAJA();
                            rendicioncaja.N_VENTA = lt_RENDICION_CAJA.GetString("N_VENTA");
                            rendicioncaja.DOC_TRIB = lt_RENDICION_CAJA.GetString("DOC_TRIB");
                            rendicioncaja.CAJERO = lt_RENDICION_CAJA.GetString("CAJERO");
                            rendicioncaja.FEC_EMI = lt_RENDICION_CAJA.GetString("FEC_EMI");
                            rendicioncaja.FEC_VENC = lt_RENDICION_CAJA.GetString("FEC_VENC");
                            rendicioncaja.MONTO = lt_RENDICION_CAJA.GetString("MONTO");
                            rendicioncaja.NAME1 = lt_RENDICION_CAJA.GetString("NAME1");
                            rendicioncaja.NUM_CHEQUE = lt_RENDICION_CAJA.GetString("NUM_CHEQUE");
                            rendicioncaja.MONTO_DIA = lt_RENDICION_CAJA.GetString("MONTO_DIA");
                            MontoChqDia = MontoChqDia + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_DIA"));
                            rendicioncaja.MONTO_FECHA = lt_RENDICION_CAJA.GetString("MONTO_FECHA");
                            MontoChqFech = MontoChqFech + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_FECHA"));
                            rendicioncaja.MONTO_EFEC = lt_RENDICION_CAJA.GetString("MONTO_EFEC");
                            MontoEfect = MontoEfect + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_EFEC"));
                            rendicioncaja.MONTO_TRANSF = lt_RENDICION_CAJA.GetString("MONTO_TRANSF");
                            MontoTransf = MontoTransf + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_TRANSF"));
                            rendicioncaja.MONTO_VALE_V = lt_RENDICION_CAJA.GetString("MONTO_VALE_V");
                            MontoValeV = MontoValeV + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_VALE_V"));
                            rendicioncaja.MONTO_DEP = lt_RENDICION_CAJA.GetString("MONTO_DEP");
                            MontoDepot = MontoDepot + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_DEP"));
                            rendicioncaja.MONTO_TARJ = lt_RENDICION_CAJA.GetString("MONTO_TARJ");
                            MontoTarj = MontoTarj + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_TARJ"));
                            rendicioncaja.MONTO_FINANC = lt_RENDICION_CAJA.GetString("MONTO_FINANC");
                            MontoFinanc = MontoFinanc + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_FINANC"));
                            rendicioncaja.MONTO_APP = lt_RENDICION_CAJA.GetString("MONTO_APP");
                            MontoApp = MontoApp + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_APP"));
                            rendicioncaja.MONTO_CREDITO = lt_RENDICION_CAJA.GetString("MONTO_CREDITO");
                            MontoCredit = MontoCredit + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_CREDITO"));
                            //rendicioncaja.PATENTE = lt_RENDICION_CAJA.GetString("PATENTE");
                            //rendicioncaja.MONTO_C_CURSE = lt_RENDICION_CAJA.GetString("MONTO_C_CURSE");
                            //MontoCCurse = MontoCCurse + Convert.ToDouble(lt_RENDICION_CAJA.GetString("MONTO_C_CURSE"));
                            rendicioncaja.DOC_SAP = lt_RENDICION_CAJA.GetString("DOC_SAP");
                            rendicioncaja.MONEDA = lt_RENDICION_CAJA.GetString("MONEDA");
                            rendicion_caja.Add(rendicioncaja);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    lt_RESUMEN_MENSUAL = BapiGetUser.GetTable("RESUMEN_MENSUAL");
                    try
                    {
                        for (int i = 0; i < lt_RESUMEN_MENSUAL.Count(); i++)
                        {
                            lt_RESUMEN_MENSUAL.CurrentIndex = i;
                            resumenmensual = new RESUMEN_MENSUAL();
                            resumenmensual.ID_SUCURSAL = lt_RESUMEN_MENSUAL.GetString("ID_SUCURSAL");
                            resumenmensual.ID_CAJA = lt_RESUMEN_MENSUAL.GetString("ID_CAJA");
                            resumenmensual.SUCURSAL = lt_RESUMEN_MENSUAL.GetString("SUCURSAL");
                            resumenmensual.NOM_CAJA = lt_RESUMEN_MENSUAL.GetString("NOM_CAJA");
                            resumenmensual.CAJERO = lt_RESUMEN_MENSUAL.GetString("CAJERO");
                            resumenmensual.AREA_VTAS = lt_RESUMEN_MENSUAL.GetString("AREA_VTAS");
                            resumenmensual.FLUJO_DOCS = lt_RESUMEN_MENSUAL.GetString("FLUJO_DOCS");
                            resumenmensual.TOTAL_MOV = lt_RESUMEN_MENSUAL.GetString("TOTAL_MOV");
                            resumenmensual.TOTAL_INGR = lt_RESUMEN_MENSUAL.GetString("TOTAL_INGR");
                            resumenmensual.MONTO_EFEC = lt_RESUMEN_MENSUAL.GetString("MONTO_EFEC");
                            MontoEfect = MontoEfect + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_EFEC"));
                            resumenmensual.MONTO_DIA = lt_RESUMEN_MENSUAL.GetString("MONTO_DIA");
                            MontoChqDia = MontoChqDia + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_DIA"));
                            resumenmensual.MONTO_FECHA = lt_RESUMEN_MENSUAL.GetString("MONTO_FECHA");
                            MontoChqFech = MontoChqFech + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_FECHA"));
                            resumenmensual.MONTO_TRANSF = lt_RESUMEN_MENSUAL.GetString("MONTO_TRANSF");
                            MontoTransf = MontoTransf + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_TRANSF"));
                            resumenmensual.MONTO_VALE_V = lt_RESUMEN_MENSUAL.GetString("MONTO_VALE_V");
                            MontoValeV = MontoValeV + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_VALE_V"));
                            resumenmensual.MONTO_DEP = lt_RESUMEN_MENSUAL.GetString("MONTO_DEP");
                            MontoDepot = MontoDepot + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_DEP"));
                            resumenmensual.MONTO_TARJ = lt_RESUMEN_MENSUAL.GetString("MONTO_TARJ");
                            MontoTarj = MontoTarj + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_TARJ"));
                            resumenmensual.MONTO_FINANC = lt_RESUMEN_MENSUAL.GetString("MONTO_FINANC");
                            MontoFinanc = MontoFinanc + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_FINANC"));
                            resumenmensual.MONTO_APP = lt_RESUMEN_MENSUAL.GetString("MONTO_APP");
                            MontoApp = MontoApp + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_APP"));
                            resumenmensual.MONTO_CREDITO = lt_RESUMEN_MENSUAL.GetString("MONTO_CREDITO");
                            MontoCredit = MontoCredit + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_CREDITO"));
                            //resumenmensual.MONTO_C_CURSE = lt_RESUMEN_MENSUAL.GetString("MONTO_C_CURSE");
                            //MontoCCurse = MontoCCurse + Convert.ToDouble(lt_RESUMEN_MENSUAL.GetString("MONTO_C_CURSE"));
                            resumenmensual.TOTAL_CAJERO = lt_RESUMEN_MENSUAL.GetString("TOTAL_CAJERO");
                            resumenmensual.MONEDA = lt_RESUMEN_MENSUAL.GetString("MONEDA");
                            resumen_mensual.Add(resumenmensual);
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
                            retorno = new RETORNO();
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
