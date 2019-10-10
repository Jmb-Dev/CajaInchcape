using System;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.IO;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Collections;
using System.Windows.Controls.Primitives;
using iTextSharp.text;
using iTextSharp.text.pdf;
//using CajaIndigo.PDFPageNumber;
using CajaIndigo.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndigo.AppPersistencia.Class.PartidasAbiertas;
using CajaIndigo.AppPersistencia.Class.Monitor.Estructura;
using CajaIndigo.AppPersistencia.Class.Monitor;
using CajaIndigo.AppPersistencia.Class.MatrizDePago.Estructura;
using CajaIndigo.AppPersistencia.Class.MatrizDePago;
using CajaIndigo.AppPersistencia.Class.DocumentosPagosMasivos;
using CajaIndigo.AppPersistencia.Class.DocumentosPagosMasivos.Estructura;
using CajaIndigo.AppPersistencia.Class.StatusPagosChq;
using CajaIndigo.AppPersistencia.Class.StatusPagosChq.Estructura;
using CajaIndigo.AppPersistencia.Class.PagoDocumentosIngreso.Estructura;
using CajaIndigo.AppPersistencia.Class.PagoDocumentosIngreso;
using CajaIndigo.AppPersistencia.Class.MaestroDeBancos.Estructura;
using CajaIndigo.AppPersistencia.Class.MaestroDeBancos;
using CajaIndigo.AppPersistencia.Class.MaestroFinancieras.Estructura;
using CajaIndigo.AppPersistencia.Class.MaestroFinancieras;
using CajaIndigo.AppPersistencia.Class.Login;
using CajaIndigo.AppPersistencia.Class.Login.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCaja;
using CajaIndigo.AppPersistencia.Class.ArqueoCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.ArqueoCaja;
using CajaIndigo.AppPersistencia.Class.PreCierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.PreCierreCaja;
using CajaIndigo.AppPersistencia.Class.CierreCajaDefinitvo.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCajaDefinitvo;
using CajaIndigo.AppPersistencia.Class.Anticipos;
using CajaIndigo.AppPersistencia.Class.PagoAnticipos;
using CajaIndigo.AppPersistencia.Class.PagoAnticipos.Estructura;
using CajaIndigo.AppPersistencia.Class.BusquedaAnulacion.Estructura;
using CajaIndigo.AppPersistencia.Class.BusquedaAnulacion;
using CajaIndigo.AppPersistencia.Class.BusquedaReimpresiones.Estructura;
using CajaIndigo.AppPersistencia.Class.BusquedaReimpresiones;
using CajaIndigo.AppPersistencia.Class.NotasDeCredito;
using CajaIndigo.AppPersistencia.Class.AnulacionComprobantes;
using CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.UsuariosCaja;
using CajaIndigo.AppPersistencia.Class.BloquearCaja;
using CajaIndigo.AppPersistencia.Class.ReimpresionFiscal;
using CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes;
using CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes.Estructura;
using CajaIndigo.AppPersistencia.Class.MaestroTarjetas;
using CajaIndigo.AppPersistencia.Class.MaestroTarjetas.Estructura;
using CajaIndigo.AppPersistencia.Class.RendicionCaja;
using CajaIndigo.AppPersistencia.Class.RendicionCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.RecaudacionVehiculos.Estructura;
using CajaIndigo.AppPersistencia.Class.RecaudacionVehiculos;
using CajaIndigo.AppPersistencia.Class.NotasDeCreditoCheck.Estructura;
using CajaIndigo.AppPersistencia.Class.NotasDeCreditoCheck;
using CajaIndigo.AppPersistencia.Class.NotasDeCreditoEmision;
using CajaIndigo.AppPersistencia.Class.GestionDeDepositos;
using CajaIndigo.AppPersistencia.Class.GestionDeDepositos.Estructura;
using CajaIndigo.AppPersistencia.Class.CheckUserAnulacion;
using CajaIndigo.AppPersistencia.Class.DepositoProceso;
using CajaIndigo.AppPersistencia.Class.ReportesCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.ReportesCaja;
using CajaIndigo.AppPersistencia.Class.PagosMasivosNew;
using CajaIndigo.AppPersistencia.Class.ReimpresionFiscal.Estructura;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;
using CajaIndigo;
using System.Text.RegularExpressions;
using System.Reflection;
namespace CajaIndigo.Vista.CierreCaja
{
    /// <summary>
    /// Interaction logic for CierreCaja.xaml
    /// </summary>
    public partial class CierreCaja : System.Windows.Window
    {
        List<DetalleViasPago> cheques = new List<DetalleViasPago>();
        List<VIAS_PAGO_MASIVO> chequesMasiv = new List<VIAS_PAGO_MASIVO>();
        List<T_DOCUMENTOS> detalledocs = new List<T_DOCUMENTOS>();
        List<T_DOCUMENTOS> partidaseleccionadas = new List<T_DOCUMENTOS>();
        List<T_DOCUMENTOSAUX> partidaseleccionadasaux = new List<T_DOCUMENTOSAUX>();
        List<T_DOCUMENTOS> monitorseleccionado = new List<T_DOCUMENTOS>();
        List<CAB_COMP> cabecera = new List<CAB_COMP>();
        List<DET_COMP> detalle = new List<DET_COMP>();
        List<DET_COMP> detalleaux = new List<DET_COMP>();
        List<DOCUMENTOS> docsreimpr = new List<DOCUMENTOS>();
        List<CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes.Estructura.VIAS_PAGO> viaspagreimprcompr = new List<CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes.Estructura.VIAS_PAGO>();
        List<VIAS_PAGO2> viaspagreimpr = new List<VIAS_PAGO2>();
        List<VIAS_PAGO2> viaspagreimpraux = new List<VIAS_PAGO2>();
        NotasDeCredito notasdecredito = new NotasDeCredito();
        PartidasAbiertas partidasabiertas = new PartidasAbiertas();
        LOG_APERTURA LogOpen = new LOG_APERTURA();


        Anticipos anticipos = new Anticipos();
        DocumentosPagosMasivos documentospagosmasivos = new DocumentosPagosMasivos();
        Monitor monitor = new Monitor();
        PagoDocumentosIngreso pagodocumentosingreso = new PagoDocumentosIngreso();
        PagoAnticipos pagoanticipos = new PagoAnticipos();
        MaestroBancos maestrobancos = new MaestroBancos();
        MaestroFinancieras maestrofinanc = new MaestroFinancieras();
        MaestroTarjetas maestrotarjetas = new MaestroTarjetas();
        DispatcherTimer timer = new DispatcherTimer();
        Recaudacion_vehi recauda = new Recaudacion_vehi();

        List<IT_PAGOS> viapago = new List<IT_PAGOS>();
        List<IT_PAGOS_CAB> viacab = new List<IT_PAGOS_CAB>();
        List<ACT_FPAGOS> pagos = new List<ACT_FPAGOS>();
        List<RETURN> bapi_return2 = new List<RETURN>();
        List<RETORNO_PAGODOCU> bapi_return = new List<RETORNO_PAGODOCU>();
        List<DTE_SII> objImpr = new List<DTE_SII>();

        List<INFO_SOCI> InfoSociedad = new List<INFO_SOCI>();
        Recaudacion_vehi RECAUDA = new Recaudacion_vehi();
        FormatoMonedas Formato = new FormatoMonedas();
        int suma = 0;
        int count = 1;
        int count2 = 1;

        string Rutsoc = string.Empty;
        string NombSoci = string.Empty;

        string UserCaja = string.Empty;
        string PassCaja = string.Empty;
        string IdCaja = string.Empty;
        string NombCaja = string.Empty;
        string SociedadCaja = string.Empty;
        string MonedCaja = string.Empty;
        string PaisCja = string.Empty;
        string Monto = "0";
        string IdSistema = string.Empty;
        string Instancia = string.Empty;
        string mandante = string.Empty;
        string SapRouter = string.Empty;
        string server = string.Empty;
        string idioma = string.Empty;
        string ValidaCuenta = string.Empty;
        string ValidaCuenta2 = string.Empty;

        Vista.PagoDocumento.PagoDocumento PagDocum;
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.NotaCredito.NotaCredito NotaCredit;
        Vista.Reportes.Reportes Reporte;
        CajaIndigo.MainWindow Login;

        List<LOG_APERTURA> logApertura = new List<LOG_APERTURA>();
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();

        public CierreCaja()
        {
            InitializeComponent();
        }

        public CierreCaja(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
        {
            try
            {
                InitializeComponent();

                List<string> myItemsCollection = new List<string>();
                myItemsCollection.Add(moneda);
                int test = 0;
                test = cmbMoneda.Items.Count;
                textBlock6.Content = cajaconect;
                textBlock7.Content = usuariologg;
                textBlock8.Content = sucursal;
                textBlock9.Content = usuariotemp;
                lblMonto.Content = Convert.ToString(monto);
                lblSociedad.Content = sociedad;

                //lblPais.Content = pais;
                lblPassword.Content = passlogg;
                textBlock6.Content = cajaconect;
                textBlock7.Content = usuariologg;
                textBlock8.Content = sucursal;
                textBlock9.Content = usuariotemp;
                lblMonto.Content = Convert.ToString(monto);
                lblSociedad.Content = sociedad;
                txtIdSistema.Text = IdSistema;
                txtInstancia.Text = Instancia;
                txtMandante.Text = mandante;
                txtSapRouter.Text = SapRouter;
                txtServer.Text = server;
                txtIdioma.Text = idioma;
                UsuarioCaja.Text = usuariologg;
                PassUserCaja.Text = passlogg;
                idcaja.Text = cajaconect;
                NomCaja.Text = sucursal;
                SociedCaja.Text = sociedad;
                MonedaCaja.Text = moneda;
                PaisCaja.Text = pais;
                DateTime result = DateTime.Today;
                calendario.Text = Convert.ToString(result);
                cmbMoneda.Items.Clear();
                cmbMoneda.ItemsSource = myItemsCollection;
                DGLogApertura.ItemsSource = null;
                DGLogApertura.Items.Clear();
                DGLogApertura.ItemsSource = logApertura;
                logApertura2 = logApertura;
                txtMontoApert.Text = logApertura2[0].MONTO;
                lblPais.Content = logApertura2[0].LAND;


                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);



          logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
            }
        }
        private void CargarDatos()
        {
            UserCaja = UsuarioCaja.Text;
            PassCaja = PassUserCaja.Text;
            IdCaja = idcaja.Text;
            NombCaja = NomCaja.Text;
            SociedadCaja = SociedCaja.Text;
            string MonedCaja = MonedaCaja.Text;
            PaisCja = PaisCaja.Text;
            Monto = "0";
            IdSistema = txtIdSistema.Text;
            Instancia = txtInstancia.Text;
            mandante = txtMandante.Text;
            SapRouter = txtSapRouter.Text;
            server = txtServer.Text;
            idioma = txtIdioma.Text;

            if (PagoDocumentos.IsMouseOver == true)
            {
                PagDocum = new Vista.PagoDocumento.PagoDocumento(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                PagDocum.Show();
                this.Hide();
            }
            if (EmisionNC.IsMouseOver == true)
            {
                NotaCredit = new Vista.NotaCredito.NotaCredito(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                NotaCredit.Show();
                this.Hide();
            }

            if (Anulacion.IsMouseOver == true)
            {
                Anula = new Vista.Anulacion.Anulacion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Anula.Show();
                this.Hide();
            }

            if (Reimpresion.IsMouseOver == true)
            {
                Reimp = new Reimpresion.Reimpresion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Reimp.Show();
                this.Hide();
            }

            if (RecaudacionVeh.IsMouseOver == true)
            {
                Vehi = new Vehiculos.Vehiculo(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Vehi.Show();
                this.Hide();
            }
            if (CierreCajas.IsMouseOver == true)
            {
                CierCaja = new CierreCaja(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                CierCaja.Show();
                this.Hide();
            }
            if (ReportesCaja.IsMouseOver == true)
            {
                Reporte = new Reportes.Reportes(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Reporte.Show();
                this.Hide();
            }
        }

        private bool isLoaded;
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (isLoaded)
                return;
            isLoaded = true;
        }

        private void btnRendir_Click(object sender, RoutedEventArgs e)
        {
            if (btnRendir.IsMouseOver == true)
            {
                if (count >= count2)
                {
                    cargarGrid();
                }
                count2++;
            }            
            //RFC Rendicion Caja
            RendicionCaja rendicioncaja = new RendicionCaja();
            rendicioncaja.rendicioncaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                , DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content)
                , Convert.ToString(lblSociedad.Content), logApertura2[0].ID_REGISTRO, "0000000000", "0000000000", logApertura2[0].MONEDA);

            InfoSociedad = rendicioncaja.ObjInfoSoc;

            for (int q = 0; q < InfoSociedad.Count(); q++)
            {
                Rutsoc = InfoSociedad[q].STCD1;
                NombSoci = InfoSociedad[q].BUTXT;
            }

            if (rendicioncaja.detalle_rend.Count > 0)
            {
                GBResumenCaja.Visibility = Visibility.Visible;
                GBDetEfectivo.Visibility = Visibility.Visible;
                GBDetEfectivo2.Visibility = Visibility.Visible;
                
                GBCierreCaja.Visibility = Visibility.Visible;
                GBCommentCierre.Visibility = Visibility.Visible;
                DGResumenCaja.ItemsSource = null;
                DGResumenCaja.Items.Clear();
                DGResumenCaja.ItemsSource = rendicioncaja.detalle_rend;

                double MntTotalDolares = 0;
                double MntTotalEuros = 0;

                //NUEVO---------------------------------------------------------
                List<string> listaTiposMoneda = new List<string>();
                List<string> listaValores = new List<string>();

                for (int i = 0; i < rendicioncaja.resumen_viapago.Count(); i++)
                {
                    if (!listaTiposMoneda.Contains(rendicioncaja.resumen_viapago[i].MONEDA))
                    {
                        listaTiposMoneda.Add(rendicioncaja.resumen_viapago[i].MONEDA);
                    }

                    listaTiposMoneda.Remove(logApertura2[0].MONEDA);
                }
                foreach (String x in listaTiposMoneda)
                {
                    double acumulado = 0;

                    for (int i = 0; i < rendicioncaja.resumen_viapago.Count(); i++)
                    {
                        if (rendicioncaja.resumen_viapago[i].MONEDA.Equals(x))
                        {
                            acumulado = acumulado + Double.Parse(rendicioncaja.resumen_viapago[i].MONTO);
                        }
                    }

                    listaValores.Add(acumulado.ToString());
                }

                int contador = 0;

                foreach (String x in listaTiposMoneda)
                {
                    foreach (UIElement control in prueba2.Children)
                    {                      
                        if (control is System.Windows.Controls.TextBox)
                        {
                            System.Windows.Controls.TextBox tb = (System.Windows.Controls.TextBox)control;
                            if (tb.Name.Equals(x))
                            {                             
                                tb.Text = listaValores[contador];
                                break;
                            }
                        }                    
                    }                    
                    contador++;
                }

                if (logApertura2[0].MONEDA == "CLP")
                {
                    //string Valor = Convert.ToString(rendicioncaja.MontoEfect);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //decimal ValorAux = Convert.ToDecimal(Valor);
                    //string monedachil = string.Format("{0:0,0}", ValorAux);
                    //txtTotalEfectivo.Text = monedachil;
                    txtMEfect.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoEfect));
                    //
                    //Valor = Convert.ToString(rendicioncaja.MontoChqDia);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMChqDia.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoChqDia));
                    //
                    //Valor = Convert.ToString(rendicioncaja.MontoChqFech);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMChqFech.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoChqFech));
                    //
                    //Valor = Convert.ToString(rendicioncaja.MontoTransf);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    ////
                    txtMTransf.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoTransf));
                    //Valor = Convert.ToString(rendicioncaja.MontoValeV);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMValeV.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoValeV));
                    //Valor = Convert.ToString(rendicioncaja.MontoDepot);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMDepos.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoDepot));
                    //Valor = Convert.ToString(rendicioncaja.MontoTarj);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMTarj.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoTarj));
                    //Valor = Convert.ToString(rendicioncaja.MontoFinanc);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMFinanc.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoFinanc));
                    //
                    //Valor = Convert.ToString(rendicioncaja.MontoApp);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMApp.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoApp));
                    //EGRESOS

                    //Valor = Convert.ToString(rendicioncaja.MontoCredit);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMCredit.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoCredit));
                    //Valor = Convert.ToString(rendicioncaja.MontoEgresos);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMEgresos.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoEgresos));
                    //
                    //Valor = Convert.ToString(rendicioncaja.MontoIngresos);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMIngresos.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoIngresos));
                    //
                    //Valor = Convert.ToString(rendicioncaja.MontoCCurse);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMCartaC.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.MontoCCurse));
                    //
                    //Valor = Convert.ToString(rendicioncaja.SaldoTotal);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMSaldoF.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.SaldoTotal));
                    //
                    //Valor = Convert.ToString(rendicioncaja.SaldoTotal);
                    //if (Valor.Contains("-"))
                    //{
                    //    Valor = "-" + Valor.Replace("-", "");
                    //}
                    //Valor = Valor.Replace(".", "");
                    //Valor = Valor.Replace(",", "");
                    //ValorAux = Convert.ToDecimal(Valor);
                    //monedachil = string.Format("{0:0,0}", ValorAux);
                    txtTotalCaja.Text = Formato.FormatoMoneda(Convert.ToString(rendicioncaja.SaldoTotal));
                }
                else
                {
                    //string moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoEfect));
                    txtMEfect.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoEfect));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoChqDia));
                    txtMChqDia.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoChqDia));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoChqFech));
                    txtMChqFech.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoChqFech));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoTransf));
                    txtMTransf.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoTransf));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoValeV));
                    txtMValeV.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoValeV));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoDepot));
                    txtMDepos.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoDepot));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoTarj));
                    txtMTarj.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoTarj));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoFinanc));
                    txtMFinanc.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoFinanc));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoApp));
                    txtMApp.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoApp));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoCredit));
                    txtMCredit.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoCredit));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoEgresos));
                    txtMEgresos.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoEgresos));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoIngresos));
                    txtMIngresos.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoIngresos));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoCCurse));
                    txtMCartaC.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.MontoCCurse));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.SaldoTotal));
                    txtMSaldoF.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.SaldoTotal));
                    //moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.SaldoTotal));
                    txtTotalCaja.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(rendicioncaja.SaldoTotal));
                }
                txtC1.Text = rendicioncaja.C1;
                txtC5.Text = rendicioncaja.C5;
                txtC10.Text = rendicioncaja.C10;
                txtC50.Text = rendicioncaja.C50;
                txtC100.Text = rendicioncaja.C100;
                txtC500.Text = rendicioncaja.C500;
                txtC1000.Text = rendicioncaja.C1000;
                txtC2000.Text = rendicioncaja.C2000;
                txtC5000.Text = rendicioncaja.C5000;
                txtC10000.Text = rendicioncaja.C10000;
                txtC20000.Text = rendicioncaja.C20000;
                btnInformePreCierre.IsEnabled = true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos para el informe de rendición");
            }
            if ((rendicioncaja.id_arqueo != "0000000000") & (rendicioncaja.id_cierre != "0000000000"))
            {
                txtArqueo.Text = rendicioncaja.id_arqueo;
                txtNumCierre.Text = rendicioncaja.id_cierre;
                txtDiferencia.Text = "0";
                System.Windows.Forms.MessageBox.Show("Esta caja posee un proceso de arqueo y cierre previo");
            }
            else
            {
                if (rendicioncaja.id_arqueo != "0000000000")
                {
                    btnGestionDep.IsEnabled = false;
                    txtArqueo.Text = rendicioncaja.id_arqueo;
                    txtNumCierre.Text = rendicioncaja.id_cierre;
                    txtDiferencia.Text = "0";
                    System.Windows.Forms.MessageBox.Show("Esta caja posee un proceso de arqueo previo");
                }
            }
            GC.Collect();
        }

        private void btnCerrarCaja_Click(object sender, RoutedEventArgs e)
        {
            AppPersistencia.Class.CierreCaja.CierreCaja Cerrar = new AppPersistencia.Class.CierreCaja.CierreCaja();
            Cerrar.cierreTempo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(lblPais.Content), "5000", "1000", "Probando 1", "Probando 2");
            System.Windows.Forms.MessageBox.Show(Cerrar.T_Retorno[0].MESSAGE.ToString());
            if (Cerrar.status == "S")
            {
                this.IsEnabled = false;
                this.Close();
            }
        }
        //CIERRE DE CAJA DEFINITIVO
        private void btnCierreCaja_Click(object sender, RoutedEventArgs e)
        {
            //***RFC cierre de Caja
            CierreCajaDefinitivo cierrecaja = new CierreCajaDefinitivo();
            cierrecaja.cierrecajadefinitivo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), logApertura2[0].ID_REGISTRO, txtTotalCaja.Text, txtDiferencia.Text, txtCommDif.Text, txtCommCierre.Text, logApertura2[0].MONTO
                , "C", txtArqueo.Text);

            if (cierrecaja.errormessage != "")
            {
                System.Windows.Forms.MessageBox.Show(cierrecaja.errormessage);
            }
            if (cierrecaja.message != "")
            {
                System.Windows.Forms.MessageBox.Show(cierrecaja.message);
            }
            if (cierrecaja.numerocierre != "0000000000")
            {
                System.Windows.Forms.MessageBox.Show("Cierre de caja, número: " + cierrecaja.numerocierre);
                txtNumCierre.Text = cierrecaja.numerocierre;
                Inicio.IsEnabled = false;
                PagoDocumentos.IsEnabled = false;
                EmisionNC.IsEnabled = false;
                Anulacion.IsEnabled = false;
                Reimpresion.IsEnabled = false;
                RecaudacionVeh.IsEnabled = false;
                BloquearCaja.IsEnabled = false;
                btnAnulCierre.IsEnabled = true;
                btnGestionDep.IsEnabled = true;
            }
            GC.Collect();
        }

        private void ChkDepositos_Checked(object sender, RoutedEventArgs e)
        {
            List<VIAS_PAGOGDAUX> partidaseleccionadasaux2 = new List<VIAS_PAGOGDAUX>();
            partidaseleccionadasaux2.Clear();
            if (this.DGViasPagoGD.Items.Count > 0)
            {
                for (int i = 0; i < DGViasPagoGD.Items.Count - 1; i++)
                {
                    if (i == 0)
                        DGViasPagoGD.Items.MoveCurrentToFirst();
                    {
                        partidaseleccionadasaux2.Add(DGViasPagoGD.Items.CurrentItem as VIAS_PAGOGDAUX);
                    }
                    DGViasPagoGD.Items.MoveCurrentToNext();
                }
            }
            for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
            {
                partidaseleccionadasaux2[k].ISSELECTED = true;
            }
            DGViasPagoGD.ItemsSource = null;
            DGViasPagoGD.Items.Clear();
            DGViasPagoGD.ItemsSource = partidaseleccionadasaux2;
            GC.Collect();
        }

        private void ChkDepositos_Unchecked(object sender, RoutedEventArgs e)
        {
            List<VIAS_PAGOGDAUX> partidaseleccionadasaux2 = new List<VIAS_PAGOGDAUX>();
            partidaseleccionadasaux2.Clear();
            if (this.DGViasPagoGD.Items.Count > 0)
            {
                for (int i = 0; i < DGViasPagoGD.Items.Count - 1; i++)
                {
                    if (i == 0)
                        DGViasPagoGD.Items.MoveCurrentToFirst();
                    {
                        partidaseleccionadasaux2.Add(DGViasPagoGD.Items.CurrentItem as VIAS_PAGOGDAUX);
                    }
                    DGViasPagoGD.Items.MoveCurrentToNext();
                }
            }
            for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
            {
                partidaseleccionadasaux2[k].ISSELECTED = false;
            }
            DGViasPagoGD.ItemsSource = null;
            DGViasPagoGD.Items.Clear();
            DGViasPagoGD.ItemsSource = partidaseleccionadasaux2;
            GC.Collect();
        }

        private void btnArqueo_Click(object sender, RoutedEventArgs e)
        {
            List<DETALLE_ARQUEO> DetalleEfectivo = new List<DETALLE_ARQUEO>();
            DetalleEfectivo = ListaEfectivo();
            ArqueoCaja arqueoCaja = new ArqueoCaja();
            arqueoCaja.arqueocaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), logApertura2[0].MONEDA, logApertura2[0].ID_REGISTRO, "0000000000", "A", "0000000000", logApertura2[0].MONTO, DetalleEfectivo);
            if (arqueoCaja.errormessage != "")
            {
                System.Windows.MessageBox.Show(arqueoCaja.errormessage);
            }
            if (arqueoCaja.id_arqueo != "0000000000")
            {
                System.Windows.MessageBox.Show(arqueoCaja.message + "\n" + arqueoCaja.id_arqueo);
                btnCierreCaja.IsEnabled = true;
                txtArqueo.Text = arqueoCaja.id_arqueo;
            }
            else
            {
                btnArqueo.IsEnabled = false;
            }
            if (Convert.ToDouble(arqueoCaja.diferencia) != 0)
            {
                System.Windows.MessageBox.Show(arqueoCaja.message);
                label77.Visibility = Visibility.Visible;
                txtCommDif.Visibility = Visibility.Visible;
                btnArqueo.IsEnabled = false;
            }
            FormatoMonedas FM = new FormatoMonedas();
            if (logApertura2[0].MONEDA == "CLP")
            {
                string Formateo = FM.FormatoMoneda(Convert.ToString(arqueoCaja.diferencia));
                txtDiferencia.Text = Formateo;
            }
            else
            {
                string Formateo = FM.FormatoMonedaExtranjera(Convert.ToString(arqueoCaja.diferencia));
                txtDiferencia.Text = Formateo;
            }
            GC.Collect();
        }

        private void btnCalcArqueo_Click(object sender, RoutedEventArgs e)
        {
            List<DETALLE_ARQUEO> DetalleEfectivo = new List<DETALLE_ARQUEO>();
            DetalleEfectivo = ListaEfectivo();

            ArqueoCaja arqueoCaja = new ArqueoCaja();
            arqueoCaja.arqueocaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), logApertura2[0].MONEDA, logApertura2[0].ID_REGISTRO, "0000000000", "", "0000000000", logApertura2[0].MONTO, DetalleEfectivo);

            if (Convert.ToDouble(arqueoCaja.diferencia) != 0)
            {
                System.Windows.MessageBox.Show(arqueoCaja.message);
                label77.Visibility = Visibility.Visible;
                txtCommDif.Visibility = Visibility.Visible;
                btnArqueo.IsEnabled = false;
            }
            else
            {
                btnArqueo.IsEnabled = true;
            }
            if (logApertura2[0].MONEDA == "CLP")
            {
                //string Valor = Convert.ToString(arqueoCaja.diferencia);
                //if (Valor.Contains("-"))
                //{
                //    Valor = "-" + Valor.Replace("-", "");
                //}
                //Valor = Valor.Replace(".", "");
                //Valor = Valor.Replace(",", "");
                //decimal ValorAux = Convert.ToDecimal(Valor);
                //string monedachil = string.Format("{0:0,0}", ValorAux);
                txtDiferencia.Text = Formato.FormatoMoneda(Convert.ToString(arqueoCaja.diferencia));
            }
            else
            {
                //string moneda = string.Format("{0:0,0.##}", Convert.ToString(arqueoCaja.diferencia));
                txtDiferencia.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(arqueoCaja.diferencia));
            }
            GC.Collect();
        }

        private void btnGestionDep_Click(object sender, RoutedEventArgs e)
        {
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            DGViasPagoGD.ItemsSource = null;
            DGViasPagoGD.Items.Clear();
            cmbBancoDest.ItemsSource = null;
            cmbBancoDest.Items.Clear();
            cmbCuentaContable.ItemsSource = null;
            cmbCuentaContable.Items.Clear();
            GestionDeDepositos gdepot = new GestionDeDepositos();
            if (txtArqueo.Text != "")
            {
                gdepot.gestiondedepositos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                    , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                    , Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content), logApertura2[0].ID_REGISTRO, txtNumCierre.Text, txtArqueo.Text);
                if (gdepot.errormessage != "")
                {
                    System.Windows.Forms.MessageBox.Show(gdepot.errormessage);
                }
                if (gdepot.message != "")
                {
                    System.Windows.Forms.MessageBox.Show(gdepot.message);
                }
                if (gdepot.vpgestiondepositos.Count > 0)
                {
                    GBGestionDeBancos.Visibility = Visibility.Visible;

                    List<VIAS_PAGOGDAUX> partidaopen = new List<VIAS_PAGOGDAUX>();

                    for (int k = 0; k < gdepot.vpgestiondepositos.Count; k++)
                    {
                        VIAS_PAGOGDAUX partOpen = new VIAS_PAGOGDAUX();
                        partOpen.ISSELECTED = false;
                        partOpen.BANCO = gdepot.vpgestiondepositos[k].BANCO;
                        partOpen.BELNR = gdepot.vpgestiondepositos[k].BELNR;
                        partOpen.BELNR_DEP = gdepot.vpgestiondepositos[k].BELNR_DEP;
                        partOpen.CTA_BANCO = gdepot.vpgestiondepositos[k].CTA_BANCO;
                        partOpen.FEC_DEPOSITO = gdepot.vpgestiondepositos[k].FEC_DEPOSITO;
                        partOpen.FECHA_EMISION = gdepot.vpgestiondepositos[k].FECHA_EMISION;
                        partOpen.FECHA_VENC = gdepot.vpgestiondepositos[k].FECHA_VENC;
                        partOpen.HKONT = gdepot.vpgestiondepositos[k].HKONT;
                        partOpen.ID_APERTURA = gdepot.vpgestiondepositos[k].ID_APERTURA;
                        partOpen.ID_BANCO = gdepot.vpgestiondepositos[k].ID_BANCO;
                        partOpen.ID_CAJA = gdepot.vpgestiondepositos[k].ID_CAJA;
                        partOpen.ID_CIERRE = gdepot.vpgestiondepositos[k].ID_CIERRE;
                        partOpen.ID_COMPROBANTE = gdepot.vpgestiondepositos[k].ID_COMPROBANTE;
                        partOpen.ID_DEPOSITO = gdepot.vpgestiondepositos[k].ID_DEPOSITO;
                        partOpen.ID_DETALLE = gdepot.vpgestiondepositos[k].ID_DETALLE;
                        partOpen.MONEDA = gdepot.vpgestiondepositos[k].MONEDA;
                        partOpen.MONTO_DOC = gdepot.vpgestiondepositos[k].MONTO_DOC;
                        partOpen.NUM_DEPOSITO = gdepot.vpgestiondepositos[k].NUM_DEPOSITO;
                        partOpen.NUM_DOC = gdepot.vpgestiondepositos[k].NUM_DOC;
                        partOpen.SELECCION = gdepot.vpgestiondepositos[k].SELECCION;
                        partOpen.SOCIEDAD = gdepot.vpgestiondepositos[k].SOCIEDAD;
                        partOpen.TEXT_BANCO = gdepot.vpgestiondepositos[k].TEXT_BANCO;
                        partOpen.TEXT_VIA_PAGO = gdepot.vpgestiondepositos[k].TEXT_VIA_PAGO;
                        partOpen.USUARIO = gdepot.vpgestiondepositos[k].USUARIO;
                        partOpen.VIA_PAGO = gdepot.vpgestiondepositos[k].VIA_PAGO;
                        partOpen.ZUONR = gdepot.vpgestiondepositos[k].ZUONR;
                        partidaopen.Add(partOpen);
                    }
                    DGViasPagoGD.ItemsSource = partidaopen;
                    List<string> BancoDst = new List<string>();
                    List<string> CtaContable = new List<string>();
                    for (int i = 0; i < gdepot.BancoDeposito.Count; i++)
                    {
                        BancoDst.Add(gdepot.BancoDeposito[i].BANKN + "-" + gdepot.BancoDeposito[i].BANKL + "-" + gdepot.BancoDeposito[i].BANKA + "-" + gdepot.BancoDeposito[i].WAERS);
                        CtaContable.Add(gdepot.BancoDeposito[i].HKONT);
                    }
                    cmbBancoDest.ItemsSource = BancoDst;
                    cmbCuentaContable.ItemsSource = CtaContable;
                }
            }
        }

        private void btnAnulCierre_Click(object sender, RoutedEventArgs e)
        {
            CierreCajaDefinitivo cierrecaja = new CierreCajaDefinitivo();
            cierrecaja.cierrecajadefinitivo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), logApertura2[0].ID_REGISTRO, txtTotalCaja.Text, txtDiferencia.Text, txtCommDif.Text, txtCommCierre.Text, logApertura2[0].MONTO
                , "N", txtArqueo.Text);

            if (cierrecaja.errormessage != "")
            {
                System.Windows.Forms.MessageBox.Show(cierrecaja.errormessage);
            }
            if (cierrecaja.message != "")
            {
                System.Windows.Forms.MessageBox.Show(cierrecaja.message);
                txtNumCierre.Text = "";
                txtArqueo.Text = "";
                EmisionNC.IsEnabled = true;
                Anulacion.IsEnabled = true;
                btnGestionDep.IsEnabled = false;
                btnAnulCierre.IsEnabled = false;
                btnCierreCaja.IsEnabled = false;
                btnArqueo.IsEnabled = false;
            }
            GC.Collect();
        }

        private void btnDepositos_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DepositoProceso depotProcess = new DepositoProceso();
                List<VIAS_PAGOGDAUX> Comprobante = new List<VIAS_PAGOGDAUX>();
                Comprobante.Clear();
                if (this.DGViasPagoGD.Items.Count > 0)
                {
                    for (int i = 0; i < DGViasPagoGD.Items.Count - 1; i++)
                    {
                        if (i == 0)
                        {
                            DGViasPagoGD.Items.MoveCurrentToFirst();
                        }
                        if (DGViasPagoGD.Items.CurrentItem != null)
                        {
                            Comprobante.Add(DGViasPagoGD.Items.CurrentItem as VIAS_PAGOGDAUX);
                        }
                        DGViasPagoGD.Items.MoveCurrentToNext();
                    }
                }
                List<VIAS_PAGOGD> partidaopen = new List<VIAS_PAGOGD>();
                List<VIAS_PAGOGDAUX> partidaopen2 = new List<VIAS_PAGOGDAUX>();

                for (int k = 0; k < Comprobante.Count; k++)
                {
                    if (Comprobante[k].ISSELECTED == true)
                    {
                        VIAS_PAGOGD partOpen = new VIAS_PAGOGD();
                        partOpen.BANCO = Comprobante[k].BANCO;
                        partOpen.BELNR = Comprobante[k].BELNR;
                        partOpen.BELNR_DEP = Comprobante[k].BELNR_DEP;
                        partOpen.CTA_BANCO = Comprobante[k].CTA_BANCO;
                        partOpen.FEC_DEPOSITO = Comprobante[k].FEC_DEPOSITO;
                        partOpen.FECHA_EMISION = Comprobante[k].FECHA_EMISION;
                        partOpen.FECHA_VENC = Comprobante[k].FECHA_VENC;
                        partOpen.HKONT = Comprobante[k].HKONT;
                        partOpen.ID_APERTURA = Comprobante[k].ID_APERTURA;
                        partOpen.ID_BANCO = Comprobante[k].ID_BANCO;
                        partOpen.ID_CAJA = Comprobante[k].ID_CAJA;
                        partOpen.ID_CIERRE = Comprobante[k].ID_CIERRE;
                        partOpen.ID_COMPROBANTE = Comprobante[k].ID_COMPROBANTE;
                        partOpen.ID_DEPOSITO = Comprobante[k].ID_DEPOSITO;
                        partOpen.ID_DETALLE = Comprobante[k].ID_DETALLE;
                        partOpen.MONEDA = Comprobante[k].MONEDA;
                        partOpen.MONTO_DOC = Comprobante[k].MONTO_DOC;
                        partOpen.NUM_DEPOSITO = Comprobante[k].NUM_DEPOSITO;
                        partOpen.NUM_DOC = Comprobante[k].NUM_DOC;
                        partOpen.SELECCION = Comprobante[k].SELECCION;
                        partOpen.SOCIEDAD = Comprobante[k].SOCIEDAD;
                        partOpen.TEXT_BANCO = Comprobante[k].TEXT_BANCO;
                        partOpen.TEXT_VIA_PAGO = Comprobante[k].TEXT_VIA_PAGO;
                        partOpen.USUARIO = Comprobante[k].USUARIO;
                        partOpen.VIA_PAGO = Comprobante[k].VIA_PAGO;
                        partOpen.ZUONR = Comprobante[k].ZUONR;
                        partidaopen.Add(partOpen);
                    }
                }
                bool Validador = true;

                if (partidaopen.Count > 0)
                {
                    for (int k = 0; k < partidaopen.Count; k++)
                    {
                        if (partidaopen[k].SELECCION == "P")
                        {
                            Validador = false;
                        }
                         ValidaCuenta = partidaopen[k].MONEDA;
                    }
                    string[] split = cmbBancoDest.SelectedValue.ToString().Split(new Char[] { '-' });
                    ValidaCuenta2 = split[3];
                    if (Validador == true)
                    {
                        if (ValidaCuenta == ValidaCuenta2)
                        {
                            depotProcess.depositoproceso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                                , Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content), logApertura2[0].ID_REGISTRO, txtNumCierre.Text, txtArqueo.Text, DPFechaDeposito.Text
                                , txtNumDeposito.Text, partidaopen, cmbBancoDest.Text, cmbCuentaContable.Text);

                            if (depotProcess.errormessage != "")
                            {
                                System.Windows.Forms.MessageBox.Show(depotProcess.errormessage);
                            }
                            if (depotProcess.message != "")
                            {
                                System.Windows.Forms.MessageBox.Show(depotProcess.message);
                            }
                            if (depotProcess.vpgestiondepositos.Count > 0)
                            {
                                GBGestionDeBancos.Visibility = Visibility.Visible;
                                GestionDeDepositos gdepot = new GestionDeDepositos();

                                gdepot.gestiondedepositos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                                    , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                                    , Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content), logApertura2[0].ID_REGISTRO, txtNumCierre.Text, txtArqueo.Text);

                                if (gdepot.errormessage != "")
                                {
                                    System.Windows.Forms.MessageBox.Show(gdepot.errormessage);
                                    btnDepositos.IsEnabled = true;
                                }
                                if (gdepot.message != "")
                                {
                                    System.Windows.Forms.MessageBox.Show(gdepot.message);
                                    btnDepositos.IsEnabled = true;
                                }
                                if (gdepot.vpgestiondepositos.Count > 0)
                                {
                                    for (int k = 0; k < gdepot.vpgestiondepositos.Count; k++)
                                    {
                                        VIAS_PAGOGDAUX partOpen = new VIAS_PAGOGDAUX();
                                        partOpen.ISSELECTED = false;
                                        partOpen.BANCO = gdepot.vpgestiondepositos[k].BANCO;
                                        partOpen.BELNR = gdepot.vpgestiondepositos[k].BELNR;
                                        partOpen.BELNR_DEP = gdepot.vpgestiondepositos[k].BELNR_DEP;
                                        partOpen.CTA_BANCO = gdepot.vpgestiondepositos[k].CTA_BANCO;
                                        partOpen.FEC_DEPOSITO = gdepot.vpgestiondepositos[k].FEC_DEPOSITO;
                                        partOpen.FECHA_EMISION = gdepot.vpgestiondepositos[k].FECHA_EMISION;
                                        partOpen.FECHA_VENC = gdepot.vpgestiondepositos[k].FECHA_VENC;
                                        partOpen.HKONT = gdepot.vpgestiondepositos[k].HKONT;
                                        partOpen.ID_APERTURA = gdepot.vpgestiondepositos[k].ID_APERTURA;
                                        partOpen.ID_BANCO = gdepot.vpgestiondepositos[k].ID_BANCO;
                                        partOpen.ID_CAJA = gdepot.vpgestiondepositos[k].ID_CAJA;
                                        partOpen.ID_CIERRE = gdepot.vpgestiondepositos[k].ID_CIERRE;
                                        partOpen.ID_COMPROBANTE = gdepot.vpgestiondepositos[k].ID_COMPROBANTE;
                                        partOpen.ID_DEPOSITO = gdepot.vpgestiondepositos[k].ID_DEPOSITO;
                                        partOpen.ID_DETALLE = gdepot.vpgestiondepositos[k].ID_DETALLE;
                                        partOpen.MONEDA = gdepot.vpgestiondepositos[k].MONEDA;
                                        partOpen.MONTO_DOC = gdepot.vpgestiondepositos[k].MONTO_DOC;
                                        partOpen.NUM_DEPOSITO = gdepot.vpgestiondepositos[k].NUM_DEPOSITO;
                                        partOpen.NUM_DOC = gdepot.vpgestiondepositos[k].NUM_DOC;
                                        partOpen.SELECCION = gdepot.vpgestiondepositos[k].SELECCION;
                                        partOpen.SOCIEDAD = gdepot.vpgestiondepositos[k].SOCIEDAD;
                                        partOpen.TEXT_BANCO = gdepot.vpgestiondepositos[k].TEXT_BANCO;
                                        partOpen.TEXT_VIA_PAGO = gdepot.vpgestiondepositos[k].TEXT_VIA_PAGO;
                                        partOpen.USUARIO = gdepot.vpgestiondepositos[k].USUARIO;
                                        partOpen.VIA_PAGO = gdepot.vpgestiondepositos[k].VIA_PAGO;
                                        partOpen.ZUONR = gdepot.vpgestiondepositos[k].ZUONR;
                                        partidaopen2.Add(partOpen);
                                    }
                                    GBGestionDeBancos.Visibility = Visibility.Visible;
                                    DGViasPagoGD.ItemsSource = partidaopen2;
                                    btnDepositos.IsEnabled = true;
                                    btnSalirCaja.IsEnabled = true;
                                }
                            }
                            List<string> BancoDst = new List<string>();
                            try
                            {
                                DPFechaDeposito.Text = "";
                                cmbCuentaContable.Text = "";
                                cmbBancoDest.Text = "";
                                txtNumDeposito.Text = "";
                            }
                            catch (Exception ex)
                            {
                                Console.Write(ex.Message, ex.StackTrace);
                            }
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("Cuenta No es correcta para deposito, Seleccione cuenta correcta");
                        }
                    }
                    else
                    {
                        DGViasPagoGD.UnselectAll();
                        System.Windows.Forms.MessageBox.Show("Registro(s) ya procesado(s).");
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Registros no seleccionados. Seleccione los registros a depositar");
                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
            }
            GC.Collect();
        }

        private void btnSalirCaja_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            CajaIndigo.MainWindow Login;
            Login = new MainWindow();
            Login.Show();          
        }

        private void btnInformePreCierre_Click(object sender, RoutedEventArgs e)
        {
            ImpresionInformePreCierre();
        }

        private void ImpresionInformePreCierre()
        {
            try
            {
                Document pdfcommande = new Document(PageSize.LETTER.Rotate(), 20f, 20f, 20f, 20f); 
                string direct = Convert.ToString(System.IO.Path.GetTempPath());
                string fecha = Convert.ToString(DateTime.Today);
                fecha = fecha.Replace("/", "");
                fecha = fecha.Substring(0, 8);
                direct = "PrecierreCaja" + fecha + ".pdf";
                txtDirect.Text = direct;
                try
                {
                    PdfWriter.GetInstance(pdfcommande, new FileStream(direct, FileMode.OpenOrCreate));
                    pdfcommande.Open();

                    PdfPTable table2 = new PdfPTable(DGResumenCaja.Columns.Count);
                    table2.TotalWidth = 750f;
                    table2.LockedWidth = true;
                    table2.HeaderRows = 2;
                    table2.SpacingAfter = 30f;

                    List<DETALLE_REND> ViasPago = new List<DETALLE_REND>();
                    for (int k = 0; k < DGResumenCaja.Items.Count; k++)
                    {
                        if (k == 0)
                        {
                            DGResumenCaja.Items.MoveCurrentToFirst();
                        }
                        ViasPago.Add(DGResumenCaja.Items.CurrentItem as DETALLE_REND);
                        DGResumenCaja.Items.MoveCurrentToNext();
                    }

                    PdfPCell cell2 = new PdfPCell(new Phrase(Convert.ToString(label11.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    cell2.Padding = 10f;
                    cell2.Colspan = DGResumenCaja.Columns.Count;
                    cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                    table2.AddCell(cell2);
                    foreach (DataGridColumn column in DGResumenCaja.Columns)
                    {
                        table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    }
                    for (int k = 0; k < ViasPago.Count; k++)
                    {
                        if (ViasPago[k] != null)
                        {
                            try
                            {
                                PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].N_VENTA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow1);
                                PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FEC_EMI), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow2);
                                PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FEC_VENC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow3);
                                PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow4);
                                PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NAME1), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow5);
                                PdfPCell cellrow6 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_EFEC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow6);
                                PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONEDA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow7);
                                PdfPCell cellrow8 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NUM_CHEQUE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow8);
                                PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_DIA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow9);
                                PdfPCell cellrow10 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_FECHA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow10);
                                PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_TRANSF), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow11);
                                PdfPCell cellrow12 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_VALE_V), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow12);
                                PdfPCell cellrow13 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_DEP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow13);
                                PdfPCell cellrow14 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_TARJ), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow14);
                                PdfPCell cellrow15 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_FINANC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow15);
                                PdfPCell cellrow16 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_APP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow16);
                                PdfPCell cellrow17 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_CREDITO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                cellrow6.Left = 2f;
                                table2.AddCell(cellrow17);
                                PdfPCell cellrow18 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_C_CURSE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow18);
                                PdfPCell cellrow19 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].PATENTE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                cellrow8.Left = 5f;
                                table2.AddCell(cellrow19);
                                PdfPCell cellrow20 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DOC_SAP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow20);
                            }
                            catch (Exception ex)
                            {
                                Console.Write(ex.Message, ex.StackTrace);
                            }
                        }
                    }
                    pdfcommande.Add(iTextSharp.text.PageSize.LETTER);

                    string texto = Rutsoc;
                    string texto1 = NombSoci;
                    while (texto.Length < 30)
                    {
                        texto = texto + " ";
                    }
                    iTextSharp.text.Chunk itxtSoci = new Chunk(texto, FontFactory.GetFont("ARIAL", 13, iTextSharp.text.Font.BOLDITALIC));
                    pdfcommande.Add(itxtSoci);
                    iTextSharp.text.Chunk itxtSpace1 = new Chunk("                                                                                 ", FontFactory.GetFont("ARIAL", 13, iTextSharp.text.Font.BOLDITALIC));
                    pdfcommande.Add(itxtSpace1);

                    //Datos Fecha
                    texto = "";
                    DateTime FechaAct = DateTime.Now;
                    texto = "Fecha: " + Convert.ToString(FechaAct.ToString("dd-MM-yyyy"));
                    iTextSharp.text.Chunk itxtFecha = new iTextSharp.text.Chunk(texto);
                    itxtFecha.Font.Size = 10;
                    pdfcommande.Add(itxtFecha);
                    texto = "";
                    pdfcommande.Add(new iTextSharp.text.Paragraph(texto));
                    //Datos RUT Sociedad 

                    while (texto1.Length < 30)
                    {
                        texto1 = texto1 + " ";
                    }
                    iTextSharp.text.Chunk itxtRUTSoci = new Chunk(texto1, FontFactory.GetFont("ARIAL", 13, iTextSharp.text.Font.BOLDITALIC));
                    pdfcommande.Add(itxtRUTSoci);
                    iTextSharp.text.Chunk itxtSpace2 = new Chunk("                                                                                         ", FontFactory.GetFont("ARIAL", 13, iTextSharp.text.Font.BOLDITALIC));
                    pdfcommande.Add(itxtSpace2);
                    //Datos hora
                    DateTime hora = DateTime.Now;
                    texto = "Hora: " + Convert.ToString(hora.ToString("HH:mm"));
                    iTextSharp.text.Chunk itxtHora = new iTextSharp.text.Chunk(texto);
                    itxtHora.Font.Size = 10;
                    pdfcommande.Add(itxtHora);
                    texto = "";
                    pdfcommande.Add(new iTextSharp.text.Paragraph(texto));

                    texto = Convert.ToString(lblSociedad.Content);
                    while (texto.Length < 30)
                    {
                        texto = texto + " ";
                    }
                    iTextSharp.text.Chunk itxtSociSoci = new Chunk(texto, FontFactory.GetFont("ARIAL", 13, iTextSharp.text.Font.BOLDITALIC));
                    pdfcommande.Add(itxtSociSoci);
                    iTextSharp.text.Chunk itxtSpace3 = new Chunk("                                                                                              ", FontFactory.GetFont("ARIAL", 13, iTextSharp.text.Font.BOLDITALIC));
                    pdfcommande.Add(itxtSpace3);
                    //Datos usuario
                    texto = "Usuario:" + Convert.ToString(textBlock7.Content);
                    iTextSharp.text.Chunk itxtUsuario = new iTextSharp.text.Chunk(texto);
                    //itxtUsuario.IndentationLeft = 600;
                    itxtUsuario.Font.Size = 10;
                    pdfcommande.Add(itxtUsuario);
                    texto = "";
                    pdfcommande.Add(new iTextSharp.text.Paragraph(texto));

                    //Titulo
                    texto = "Informe de rendición de caja";
                    iTextSharp.text.Paragraph itxtTitulo = new iTextSharp.text.Paragraph(texto);
                    itxtTitulo.IndentationLeft = 300;
                    itxtTitulo.Font.Size = 14;
                    itxtTitulo.SpacingBefore = 30f;
                    itxtTitulo.SpacingAfter = 5f;
                    pdfcommande.Add(itxtTitulo);
                    //Fecha desde hasta del Informe
                    DateTime fecdesde = DateTime.Now;

                    texto = "Desde:" + Convert.ToString(fecdesde.ToString("dd-MM-yyyy") + " Hasta:" + Convert.ToString(fecdesde.ToString("dd-MM-yyyy")));
                    iTextSharp.text.Paragraph itxtfechDesdeHasta = new iTextSharp.text.Paragraph(texto);
                    itxtfechDesdeHasta.IndentationLeft = 280;
                    itxtfechDesdeHasta.Font.Size = 10;
                    itxtfechDesdeHasta.SpacingAfter = 5f;
                    pdfcommande.Add(itxtfechDesdeHasta);

                    //Datos Caja
                    texto = "Id Caja:" + Convert.ToString(textBlock6.Content);
                    iTextSharp.text.Paragraph itxtIngreso = new iTextSharp.text.Paragraph(texto);
                    itxtIngreso.IndentationLeft = 15;
                    itxtIngreso.Font.Size = 10;
                    pdfcommande.Add(itxtIngreso);
                    //Datos de nota venta
                    texto = "Sucursal" + Convert.ToString(textBlock8.Content);
                    iTextSharp.text.Paragraph itxtNotaVta = new iTextSharp.text.Paragraph(texto);
                    itxtNotaVta.IndentationLeft = 15;
                    itxtNotaVta.Font.Size = 10;
                    itxtNotaVta.SpacingAfter = 10;
                    pdfcommande.Add(itxtNotaVta);
                    pdfcommande.Add(table2);
                    pdfcommande.Close();
                    pdfcommande.Dispose();
                }
                catch (Exception ex)
                {
                    Console.Write(ex.Message, ex.StackTrace);
                }

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = direct;
                proc.Start();
                proc.Close();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
            }
        }
        private void btnPreCierre_Click(object sender, RoutedEventArgs e)
        {
            if ((DPickDesde.Text != "") & (DPickHasta.Text != ""))
            {
                //RFC REPORTE DE CAJAS
                ReportesCaja reportcajas = new ReportesCaja();
                reportcajas.reportescaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                    , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                    , Convert.ToString(lblPais.Content), Convert.ToString(lblSociedad.Content), logApertura2[0].ID_REGISTRO, txtNumCierre.Text, "1");
                //RFC REPORTE DE CAJAS
                if (reportcajas.rendicion_caja.Count > 0)
                {
                    if (chkExcel.IsChecked == false)
                    {
                        ImpresionReporteCaja(reportcajas.rendicion_caja, reportcajas.resumen_mensual, reportcajas.resumen_caja, reportcajas.SociedadR, reportcajas.Empresa, reportcajas.Sucursal
                          , reportcajas.RUT, reportcajas.FechArqueo, reportcajas.FechaArqueoHasta, reportcajas.Tipo);
                    }
                    else
                    {
                        ExportaDataToExcel(reportcajas.rendicion_caja, reportcajas.resumen_mensual, reportcajas.resumen_caja, reportcajas.SociedadR, reportcajas.Empresa, reportcajas.Sucursal
                            , reportcajas.RUT, reportcajas.FechArqueo, reportcajas.FechaArqueoHasta, reportcajas.Tipo, Convert.ToString(textBlock6.Content));
                    }
                    DGRendicionCajaRep.ItemsSource = null;
                    DGRendicionCajaRep.Items.Clear();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("No existen datos para el período solicitado");
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Introduzca la fecha de inicio y fecha final del período a revisar");
            }
            GC.Collect();
        }

        private void btnResumenMovimientos_Click(object sender, RoutedEventArgs e)
        {
            if ((DPickDesde.Text != "") & (DPickHasta.Text != ""))
            {
                //RFC REPORTE DE CAJAS
                ReportesCaja reportcajas = new ReportesCaja();
                reportcajas.reportescaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                    , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                    , Convert.ToString(lblPais.Content), Convert.ToString(lblSociedad.Content), logApertura2[0].ID_REGISTRO, txtNumCierre.Text, "2");

                if (reportcajas.resumen_mensual.Count > 0)
                {
                    if (chkExcel.IsChecked == false)
                    {
                        ImpresionReporteCaja(reportcajas.rendicion_caja, reportcajas.resumen_mensual, reportcajas.resumen_caja, reportcajas.SociedadR, reportcajas.Empresa, reportcajas.Sucursal
                        , reportcajas.RUT, reportcajas.FechArqueo, reportcajas.FechaArqueoHasta, reportcajas.Tipo);
                    }
                    else
                    {
                        ExportaDataToExcel(reportcajas.rendicion_caja, reportcajas.resumen_mensual, reportcajas.resumen_caja, reportcajas.SociedadR, reportcajas.Empresa, reportcajas.Sucursal
                            , reportcajas.RUT, reportcajas.FechArqueo, reportcajas.FechaArqueoHasta, reportcajas.Tipo, Convert.ToString(textBlock6.Content));
                    }
                    DGResumenMovimientosRep.ItemsSource = null;
                    DGResumenMovimientosRep.Items.Clear();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Introduzca la fecha de inicio y fecha final del período a revisar");
            }
            GC.Collect();
        }

        private void btnResumenCajas_Click(object sender, RoutedEventArgs e)
        {
            if ((DPickDesde.Text != "") & (DPickHasta.Text != ""))
            {
                //RFC REPORTE DE CAJAS
                ReportesCaja reportcajas = new ReportesCaja();
                reportcajas.reportescaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                    , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                    , Convert.ToString(lblPais.Content), Convert.ToString(lblSociedad.Content), logApertura2[0].ID_REGISTRO, txtNumCierre.Text, "3");
                //RFC REPORTE DE CAJAS
                if (reportcajas.resumen_caja.Count > 0)
                {
                    if (chkExcel.IsChecked == false)
                    {
                        ImpresionReporteCaja(reportcajas.rendicion_caja, reportcajas.resumen_mensual, reportcajas.resumen_caja, reportcajas.SociedadR, reportcajas.Empresa, reportcajas.Sucursal
                          , reportcajas.RUT, reportcajas.FechArqueo, reportcajas.FechaArqueoHasta, reportcajas.Tipo);
                    }
                    else
                    {
                        ExportaDataToExcel(reportcajas.rendicion_caja, reportcajas.resumen_mensual, reportcajas.resumen_caja, reportcajas.SociedadR, reportcajas.Empresa, reportcajas.Sucursal
                            , reportcajas.RUT, reportcajas.FechArqueo, reportcajas.FechaArqueoHasta, reportcajas.Tipo, Convert.ToString(textBlock6.Content));
                    }
                    DGResumenCajasRep.ItemsSource = null;
                    DGResumenCajasRep.Items.Clear();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Introduzca la fecha de inicio y fecha final del período a revisar");
            }
            GC.Collect();
        }

        private void cmbBancoDest_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btnDepositos.IsEnabled = true;
        }

        //SELECCION DE CUENTA CONTABLE AUTOMATICA DE ACUERDO AL BANCO DESTINO EN GESTION DE DEPOSITOS
        private void cmbBancoDest_DropDownClosed(object sender, EventArgs e)
        {
            int posicion;
            posicion = cmbBancoDest.SelectedIndex;
            cmbCuentaContable.SelectedIndex = posicion;
            GC.Collect();
        }

        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {
            GC.Collect();
        }

        private void cargarGrid()
        {
            CajaIndigo.AppPersistencia.Class.CierreCaja.CierreCaja CierreDef = new CajaIndigo.AppPersistencia.Class.CierreCaja.CierreCaja();
            CierreDef.OtrasMonedas(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, logApertura2[0].ID_CAJA, Convert.ToString(lblPais.Content));
            List<string> listaTiposMoneda = new List<string>();
            for (int i = 0; i < CierreDef.MonedExtr.Count(); i++)
            {
                if (!listaTiposMoneda.Contains(CierreDef.MonedExtr[i].MONEDA))
                {
                    listaTiposMoneda.Add(CierreDef.MonedExtr[i].MONEDA);
                }

                listaTiposMoneda.Remove(logApertura2[0].MONEDA);
            }
            if (listaTiposMoneda.Count > 0)
            {
                prueba2.ShowGridLines = true;
                ColumnDefinition gridCol1 = new ColumnDefinition();
                ColumnDefinition gridCol2 = new ColumnDefinition();
                prueba2.ColumnDefinitions.Add(gridCol1);
                prueba2.ColumnDefinitions.Add(gridCol2);
                RowDefinition gridRow1 = new RowDefinition();
                gridRow1.Height = new GridLength(45);
                RowDefinition gridRow2 = new RowDefinition();
                gridRow2.Height = new GridLength(45);
                prueba2.RowDefinitions.Add(gridRow1);
                prueba2.RowDefinitions.Add(gridRow2);

                for (int i = 0; i < listaTiposMoneda.Count(); i++)
                {
                    // Crear Caja de Textos
                    System.Windows.Controls.TextBox MyTextBox = new System.Windows.Controls.TextBox();
                    System.Windows.Controls.TextBox MyTextBox2 = new System.Windows.Controls.TextBox();
                    MyTextBox.Text = CierreDef.MonedExtr[i].MONEDA;
                    MyTextBox.Name = CierreDef.MonedExtr[i].MONEDA + "LBL";

                    MyTextBox.FontSize = 12;
                    MyTextBox.Width = 100;
                    MyTextBox.VerticalAlignment = VerticalAlignment.Top;
                    MyTextBox.IsEnabled = false;
                    MyTextBox.FontWeight = FontWeights.Bold;
                    Grid.SetRow(MyTextBox, i);
                    Grid.SetColumn(MyTextBox, 0);
                    prueba2.Children.Add(MyTextBox);

                    MyTextBox2.Width = 100;
                    MyTextBox2.FontSize = 12;
                    MyTextBox2.Name = CierreDef.MonedExtr[i].MONEDA;
                    MyTextBox2.VerticalAlignment = VerticalAlignment.Top;
                    Grid.SetRow(MyTextBox2, i);
                    Grid.SetColumn(MyTextBox2, 1);
                    prueba2.Children.Add(MyTextBox2);
                }
            }
        }
        private void txtC1_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total1 = 0;
            bool digit = true;
            foreach (char value in txtC1.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC1.Text != "")
                {
                    Total1 = Convert.ToDouble(lbl1.Content) * Convert.ToDouble(txtC1.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT1.Text = FM.FormatoMoneda(Convert.ToString(Total1));
                    }
                    else
                    {
                        txtT1.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total1));
                    }
                }
                else
                    txtC1.Text = "0";

            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");

            SumaEfectivoPorDenominacion();
            GC.Collect();
        }

        private void txtC5_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total5 = 0;
            bool digit = true;
            foreach (char value in txtC5.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC5.Text != "")
                {
                    Total5 = Convert.ToDouble(lbl5.Content) * Convert.ToDouble(txtC5.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT5.Text = FM.FormatoMoneda(Convert.ToString(Total5));
                    }
                    else
                    {
                        txtT5.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total5));
                    }
                }
                else
                    txtC5.Text = "0";
            }
            else
            {
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                txtC5.Text = "";
            }
            SumaEfectivoPorDenominacion();
            GC.Collect();
        }

        private void txtC10_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total10 = 0;
            bool digit = true;
            foreach (char value in txtC10.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC10.Text != "")
                {
                    Total10 = Convert.ToDouble(lbl10.Content) * Convert.ToDouble(txtC10.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT10.Text = FM.FormatoMoneda(Convert.ToString(Total10));
                    }
                    else
                    {
                        txtT10.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total10));
                    }
                }
                else
                    txtC10.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
                GC.Collect();
        }

        private void txtC50_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total50 = 0;
            bool digit = true;
            foreach (char value in txtC50.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC50.Text != "")
                {
                    Total50 = Convert.ToDouble(lbl50.Content) * Convert.ToDouble(txtC50.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT50.Text = FM.FormatoMoneda(Convert.ToString(Total50));
                    }
                    else
                    {
                        txtT50.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total50));
                    }
                }
                else
                    txtC50.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
                GC.Collect();
        }

        private void txtC100_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total100 = 0;
            bool digit = true;
            foreach (char value in txtC100.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC100.Text != "")
                {
                    Total100 = Convert.ToDouble(lbl100.Content) * Convert.ToDouble(txtC100.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT100.Text = FM.FormatoMoneda(Convert.ToString(Total100));
                    }
                    else
                    {
                        txtT100.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total100));
                    }
                    txtT100.Text = Convert.ToString(Total100);
                }
                else
                    txtC100.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
                GC.Collect();
        }

        private void txtC500_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total500 = 0;
            bool digit = true;
            foreach (char value in txtC500.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC500.Text != "")
                {
                    Total500 = Convert.ToDouble(lbl500.Content) * Convert.ToDouble(txtC500.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT500.Text = FM.FormatoMoneda(Convert.ToString(Total500));
                    }
                    else
                    {
                        txtT500.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total500));
                    }
                }
                else
                    txtC500.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
                GC.Collect();
        }

        private void txtC1000_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total1000 = 0;
            bool digit = true;
            foreach (char value in txtC1000.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC1000.Text != "")
                {
                    Total1000 = Convert.ToDouble(lbl1000.Content) * Convert.ToDouble(txtC1000.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT1000.Text = FM.FormatoMoneda(Convert.ToString(Total1000));
                    }
                    else
                    {
                        txtT1000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total1000));
                    }
                }
                else
                    txtC1000.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
                GC.Collect();
        }

        private void txtC2000_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total2000 = 0;
            bool digit = true;
            foreach (char value in txtC2000.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC2000.Text != "")
                {
                    Total2000 = Convert.ToDouble(lbl2000.Content) * Convert.ToDouble(txtC2000.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT2000.Text = FM.FormatoMoneda(Convert.ToString(Total2000));
                    }
                    else
                    {
                        txtT2000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total2000));
                    }
                }
                else
                    txtC2000.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
                GC.Collect();
        }

        private void txtC5000_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total5000 = 0;
            bool digit = true;
            foreach (char value in txtC5000.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC5000.Text != "")
                {
                    Total5000 = Convert.ToDouble(lbl5000.Content) * Convert.ToDouble(txtC5000.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT5000.Text = FM.FormatoMoneda(Convert.ToString(Total5000));
                    }
                    else
                    {
                        txtT5000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total5000));
                    }
                }
                else
                    txtC5000.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
            GC.Collect();
        }

        private void txtC10000_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total10000 = 0;
            bool digit = true;
            foreach (char value in txtC10000.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC10000.Text != "")
                {
                    Total10000 = Convert.ToDouble(lbl10000.Content) * Convert.ToDouble(txtC10000.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT10000.Text = FM.FormatoMonedaChilena(Convert.ToString(Total10000), "1");
                    }
                    else
                    {
                        txtT10000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total10000));
                    }
                    //txtT10000.Text = Convert.ToString(Total10000);
                }
                else
                    txtC10000.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                SumaEfectivoPorDenominacion();
             GC.Collect();
        }

        private void txtC20000_TextChanged(object sender, TextChangedEventArgs e)
        {
            FormatoMonedas FM = new FormatoMonedas();
            double Total20000 = 0;
            bool digit = true;
            foreach (char value in txtC20000.Text)
            {
                digit = char.IsDigit(value);
                if (digit == false)
                    break;
            }

            if (digit)
            {
                if (txtC20000.Text != "")
                {
                    Total20000 = Convert.ToDouble(lbl20000.Content) * Convert.ToDouble(txtC20000.Text);
                    if (logApertura2[0].MONEDA == "CLP")
                    {
                        txtT20000.Text = FM.FormatoMoneda(Convert.ToString(Total20000));
                    }
                    else
                    {
                        txtT20000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total20000));
                    }
                }
                else
                    txtC20000.Text = "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");

            SumaEfectivoPorDenominacion();
            GC.Collect();
        }


        private void SumaEfectivoPorDenominacion()
        {
            double Total1 = 0;
            double Total5 = 0;
            double Total10 = 0;
            double Total50 = 0;
            double Total100 = 0;
            double Total500 = 0;
            double Total1000 = 0;
            double Total2000 = 0;
            double Total5000 = 0;
            double Total10000 = 0;
            double Total20000 = 0;
            FormatoMonedas FM = new FormatoMonedas();

            if (txtT1.Text != "")
            {
                Total1 = Convert.ToDouble(txtT1.Text);
            }
            if (txtT5.Text != "")
            {
                Total5 = Convert.ToDouble(txtT5.Text);
            }

            if (txtT10.Text != "")
            {
                Total10 = Convert.ToDouble(txtT10.Text);
            }

            if (txtT50.Text != "")
            {
                Total50 = Convert.ToDouble(txtT50.Text);
            }

            if (txtT100.Text != "")
            {
                Total100 = Convert.ToDouble(txtT100.Text);
            }

            if (txtT500.Text != "")
            {
                Total500 = Convert.ToDouble(txtT500.Text);
            }

            if (txtT1000.Text != "")
            {
                Total1000 = Convert.ToDouble(txtT1000.Text);
            }
            if (txtT2000.Text != "")
            {
                Total2000 = Convert.ToDouble(txtT2000.Text);
            }

            if (txtT5000.Text != "")
            {
                Total5000 = Convert.ToDouble(txtT5000.Text);
            }

            if (txtT10000.Text != "")
            {
                Total10000 = Convert.ToDouble(txtT10000.Text);
            }

            if (txtT20000.Text != "")
            {

                Total20000 = Convert.ToDouble(txtT20000.Text);
            }

            if(logApertura2[0].MONEDA == "CLP")
            {
                txtTotalEfectivo.Text = FM.FormatoMonedaChilena(Convert.ToString(Total1 + Total5 + Total10 + Total50 + Total100 + Total500 + Total1000 + Total2000 + Total5000 + Total10000 + Total20000), "1");
            }
            else
            {
                txtTotalEfectivo.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total1 + Total5 + Total10 + Total50 + Total100 + Total500 + Total1000 + Total2000 + Total5000 + Total10000 + Total20000));
            }
            GC.Collect();
        }

        private void ExportaDataToExcel(List<RENDICION_CAJA> ListRendicionCaja, List<RESUMEN_MENSUAL> ListResumenMensual, List<RESUMEN_CAJA> ListResumenCaja, string SociedadR, string Empresa, string Sucursal
        , string RUT, string FechaDesde, string FechaHasta, string Tipo, string Caja)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                try
                {;
                    xlWorkSheet.Range["A3"].Value = Empresa;
                    xlWorkSheet.Range["A4"].Value = RUT;
                    xlWorkSheet.Range["A5"].Value = SociedadR;

                    xlWorkSheet.Range["A7"].Value = "Id Caja:";
                    xlWorkSheet.Range["B7"].Value = Caja;
                    xlWorkSheet.Range["A8"].Value = "Sucursal";
                    xlWorkSheet.Range["B8"].Value = Convert.ToString(textBlock8.Content);

                    xlWorkSheet.Range["O3"].Value = "Fecha:";
                    string FechaAFormatear = Convert.ToString(DateTime.Now).Substring(0, 10);
                    xlWorkSheet.Range["P3"].Value = String.Format("{0:dd/MM/yyyy}", FechaAFormatear);
                    xlWorkSheet.Range["O4"].Value = "Hora:";
                    xlWorkSheet.Range["P4"].Value = Convert.ToString(DateTime.Now.Hour) + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second);
                    xlWorkSheet.Range["O5"].Value = "Usuario:";
                    xlWorkSheet.Range["P5"].Value = Convert.ToString(textBlock7.Content);
                    xlWorkSheet.Range["O6"].Value = "Cajero:";
                    xlWorkSheet.Range["P6"].Value = Convert.ToString(textBlock7.Content);
                }
                catch (Exception ex)
                {
                    Console.Write(ex.Message, ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                }
                if (Tipo == "1")
                {
                    xlWorkSheet.Range["G5"].Value = "Informe de rendición de caja";
                }
                if (Tipo == "2")
                {
                    xlWorkSheet.Range["G5"].Value = "Informe resumen mensual de movimientos";
                }
                if (Tipo == "3")
                {
                    xlWorkSheet.Range["G5"].Value = "Resumen caja recaudadora";
                }

                xlWorkSheet.Range["A7"].Value = "Desde:";
                xlWorkSheet.Range["B7"].Value = FechaDesde;
                xlWorkSheet.Range["A8"].Value = "Hasta:";
                xlWorkSheet.Range["B8"].Value = FechaHasta;


                int i = 0;
                int j = 0;

                DGRendicionCajaRep.ItemsSource = ListRendicionCaja;
                DGResumenCajasRep.ItemsSource = ListResumenCaja;
                DGResumenMovimientosRep.ItemsSource = ListResumenMensual;

                double MONTO = 0;
                double TOTAL_MOV = 0;
                double TOTAL_INGR = 0;
                double MONTO_EFEC = 0;
                double MONTO_DIA = 0;
                double MONTO_FECHA = 0;
                double MONTO_TRANSF = 0;
                double MONTO_VALE_V = 0;
                double MONTO_DEP = 0;
                double MONTO_TARJ = 0;
                double MONTO_FINANC = 0;
                double MONTO_APP = 0;
                double MONTO_CREDITO = 0;
                double TOTAL_CAJERO = 0;
                if (Tipo == "1")
                {
                    try
                    {
                        List<RENDICION_CAJA> ViasPago = new List<RENDICION_CAJA>();
                        for (int k = 0; k < DGRendicionCajaRep.Items.Count; k++)
                        {

                            if (k == 0)
                            {
                                DGRendicionCajaRep.Items.MoveCurrentToFirst();
                            }
                            if (DGRendicionCajaRep.Items.CurrentItem != null)
                            {
                                ViasPago.Add(DGRendicionCajaRep.Items.CurrentItem as RENDICION_CAJA);
                            }
                            DGRendicionCajaRep.Items.MoveCurrentToNext();
                        }
                        //NOMBRES DE COLUMNAS
                        xlWorkSheet.Range["A10"].Value = "Tipo documento";
                        xlWorkSheet.Range["B10"].Value = "N° Doc. Tributario";
                        xlWorkSheet.Range["C10"].Value = "Cajero";
                        xlWorkSheet.Range["D10"].Value = "Fech. emision";
                        xlWorkSheet.Range["E10"].Value = "Fech. vencto.";
                        xlWorkSheet.Range["F10"].Value = "Monto";
                        xlWorkSheet.Range["G10"].Value = "Cliente";
                        xlWorkSheet.Range["H10"].Value = "Efectivo";
                        xlWorkSheet.Range["I10"].Value = "N° doc.";
                        xlWorkSheet.Range["J10"].Value = "Chq. al día";
                        xlWorkSheet.Range["K10"].Value = "Chq. a fecha";
                        xlWorkSheet.Range["L10"].Value = "Transferencia";
                        xlWorkSheet.Range["M10"].Value = "V. vista";
                        xlWorkSheet.Range["N10"].Value = "Depósitos";
                        xlWorkSheet.Range["O10"].Value = "Tarjetas";
                        xlWorkSheet.Range["P10"].Value = "Financiamiento";
                        xlWorkSheet.Range["Q10"].Value = "APP";
                        xlWorkSheet.Range["R10"].Value = "Crédito";
                        xlWorkSheet.Range["S10"].Value = "Doc. SAP";

                        int lineabase = 11;
                        int lineatope = lineabase + ViasPago.Count;

                        FormatoMonedas FM = new FormatoMonedas();
                        string MonedaFormateada = "";
                        for (int k = 0; k < ViasPago.Count; k++)
                        {
                            int linea = lineabase + k;
                            xlWorkSheet.Range["A" + Convert.ToString(linea)].Value = ViasPago[k].N_VENTA;
                            xlWorkSheet.Range["B" + Convert.ToString(linea)].Value = ViasPago[k].DOC_TRIB;
                            xlWorkSheet.Range["C" + Convert.ToString(linea)].Value = ViasPago[k].CAJERO;
                            xlWorkSheet.Range["D" + Convert.ToString(linea)].Value = ViasPago[k].FEC_EMI;
                            xlWorkSheet.Range["E" + Convert.ToString(linea)].Value = ViasPago[k].FEC_VENC;
                            if (k != ViasPago.Count)
                            {
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO, "2");
                                    xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = Convert.ToString(MonedaFormateada);
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO);
                                    xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO = MONTO + Convert.ToDouble(ViasPago[k].MONTO);

                                xlWorkSheet.Range["G" + Convert.ToString(linea)].Value = ViasPago[k].NAME1;

                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_EFEC, "2");
                                    xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);

                                xlWorkSheet.Range["I" + Convert.ToString(linea)].Value = ViasPago[k].NUM_CHEQUE;
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_DIA, "2");
                                    xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                    xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_FECHA, "2");
                                    xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                    xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_TRANSF, "2");
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_VALE_V, "2");
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_DEP, "2");
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_TARJ, "2");
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);

                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_FINANC, "2");
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_APP, "2");
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_CREDITO, "2");
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = ViasPago[k].DOC_SAP;
                            }
                        }
                        //TOTALES
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope)].Value = "Totales: ";
                        //Monto
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO), "1");
                            xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = Convert.ToString(MonedaFormateada);
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO));
                            xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = Convert.ToString(MONTO);
                        //Efectivo
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_EFEC), "1");
                            xlWorkSheet.Range["H" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                            xlWorkSheet.Range["H" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque al dia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_DIA), "1");
                            xlWorkSheet.Range["J" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                            xlWorkSheet.Range["J" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque a fecha
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_FECHA), "1");
                            xlWorkSheet.Range["K" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                            xlWorkSheet.Range["K" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Transferencia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_TRANSF), "1");
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Vale vista
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_VALE_V), "1");
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Deposito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_DEP), "1");
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Tarjetas
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_TARJ), "1");
                            xlWorkSheet.Range["O" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                            xlWorkSheet.Range["O" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Financiamiento
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_FINANC), "1");
                            xlWorkSheet.Range["P" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                            xlWorkSheet.Range["P" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //APP
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_APP), "1");
                            xlWorkSheet.Range["Q" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                            xlWorkSheet.Range["Q" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Credito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_CREDITO), "1");
                            xlWorkSheet.Range["R" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["R" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message, ex.StackTrace);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                }
                //Resumen Mensual de Movimientos
                if (Tipo == "2")
                {
                    try
                    {
                        List<RESUMEN_MENSUAL> ViasPago = new List<RESUMEN_MENSUAL>();
                        for (int k = 0; k < DGResumenMovimientosRep.Items.Count; k++)
                        {

                            if (k == 0)
                            {
                                DGResumenMovimientosRep.Items.MoveCurrentToFirst();
                            }
                            if (DGResumenMovimientosRep.Items.CurrentItem != null)
                            {
                                ViasPago.Add(DGResumenMovimientosRep.Items.CurrentItem as RESUMEN_MENSUAL);
                            }
                            DGResumenMovimientosRep.Items.MoveCurrentToNext();
                        }
                        //NOMBRES DE COLUMNAS
                        xlWorkSheet.Range["A10"].Value = "Id. Suc";
                        xlWorkSheet.Range["B10"].Value = "Sucursal";
                        xlWorkSheet.Range["C10"].Value = "Id. Caja";
                        xlWorkSheet.Range["D10"].Value = "Caja";
                        xlWorkSheet.Range["E10"].Value = "Cajero";
                        xlWorkSheet.Range["F10"].Value = "Area";
                        xlWorkSheet.Range["G10"].Value = "Flujo docs.";
                        xlWorkSheet.Range["H10"].Value = "Total mov.";
                        xlWorkSheet.Range["I10"].Value = "Total ingresos";
                        xlWorkSheet.Range["J10"].Value = "Efectivo";
                        xlWorkSheet.Range["K10"].Value = "Chq. al día";
                        xlWorkSheet.Range["L10"].Value = "Chq. a fecha";
                        xlWorkSheet.Range["M10"].Value = "Transferencia";
                        xlWorkSheet.Range["N10"].Value = "V. Vista";
                        xlWorkSheet.Range["O10"].Value = "Depósitos";
                        xlWorkSheet.Range["P10"].Value = "Tarjetas";
                        xlWorkSheet.Range["Q10"].Value = "Financiamiento";
                        xlWorkSheet.Range["R10"].Value = "APP";
                        xlWorkSheet.Range["S10"].Value = "Crédito";
                        xlWorkSheet.Range["T10"].Value = "Total Cajero";

                        int lineabase = 11;
                        int lineatope = lineabase + ViasPago.Count;

                        FormatoMonedas FM = new FormatoMonedas();
                        string MonedaFormateada = "";
                        for (int k = 0; k < ViasPago.Count; k++)
                        {
                            int linea = lineabase + k;
                            xlWorkSheet.Range["A" + Convert.ToString(linea)].Value = ViasPago[k].ID_SUCURSAL;
                            xlWorkSheet.Range["B" + Convert.ToString(linea)].Value = ViasPago[k].SUCURSAL;
                            xlWorkSheet.Range["C" + Convert.ToString(linea)].Value = ViasPago[k].ID_CAJA;
                            xlWorkSheet.Range["D" + Convert.ToString(linea)].Value = ViasPago[k].NOM_CAJA;
                            xlWorkSheet.Range["E" + Convert.ToString(linea)].Value = ViasPago[k].CAJERO;
                            xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = ViasPago[k].AREA_VTAS;
                            xlWorkSheet.Range["G" + Convert.ToString(linea)].Value = ViasPago[k].FLUJO_DOCS;
                            if (k != ViasPago.Count)
                            {
                                //Total Movimientos
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].TOTAL_MOV, "2");
                                    xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = Convert.ToString(MonedaFormateada);
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_MOV);
                                    xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                TOTAL_MOV = TOTAL_MOV + Convert.ToDouble(ViasPago[k].TOTAL_MOV);
                                //Total Ingresos
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].TOTAL_INGR, "2");
                                    xlWorkSheet.Range["I" + Convert.ToString(linea)].Value = Convert.ToString(MonedaFormateada);
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_INGR);
                                    xlWorkSheet.Range["I" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                TOTAL_INGR = TOTAL_INGR + Convert.ToDouble(ViasPago[k].TOTAL_INGR);
                                //Efectivo
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_EFEC, "2");
                                    xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                //Cheque al día
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_DIA, "2");
                                    xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                    xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                //Cheque a fecha
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_FECHA, "2");
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                //Transferencias
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_TRANSF, "2");
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                //Vale vistas
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_VALE_V, "2");
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                //Depositos
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_DEP, "2");
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                //Tarjetas
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_TARJ, "2");
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                //Financiamiento
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_FINANC, "2");
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                //App
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_APP, "2");
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                //Credito
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_CREDITO, "2");
                                    xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                    xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                //Total Cajero
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].TOTAL_CAJERO, "2");
                                    xlWorkSheet.Range["T" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_CAJERO);
                                    xlWorkSheet.Range["T" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                TOTAL_CAJERO = TOTAL_CAJERO + Convert.ToDouble(ViasPago[k].TOTAL_CAJERO);
                            }
                        }
                        //TOTALES
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope)].Value = "Totales: ";
                        //Total movimientos
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(TOTAL_MOV), "1");
                            xlWorkSheet.Range["H" + Convert.ToString(lineatope)].Value = Convert.ToString(MonedaFormateada);
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_MOV));
                            xlWorkSheet.Range["H" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Total ingresos
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(TOTAL_INGR), "1");
                            xlWorkSheet.Range["I" + Convert.ToString(lineatope)].Value = Convert.ToString(MonedaFormateada);
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_INGR));
                            xlWorkSheet.Range["I" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        ////Monto
                        //if (logApertura2[0].MONEDA == "CLP")
                        //{
                        //    MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO), "1");
                        //    xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = Convert.ToString(MonedaFormateada);
                        //}
                        //else
                        //{
                        //    MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO));
                        //    xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        //}
                        //xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = Convert.ToString(MONTO);
                        //Efectivo
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_EFEC), "1");
                            xlWorkSheet.Range["J" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                            xlWorkSheet.Range["J" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque al dia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_DIA), "1");
                            xlWorkSheet.Range["K" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                            xlWorkSheet.Range["K" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque a fecha
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_FECHA), "1");
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Transferencia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_TRANSF), "1");
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Vale vista
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_VALE_V), "1");
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Deposito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_DEP), "1");
                            xlWorkSheet.Range["O" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                            xlWorkSheet.Range["O" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Tarjetas
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_TARJ), "1");
                            xlWorkSheet.Range["P" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                            xlWorkSheet.Range["P" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Financiamiento
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_FINANC), "1");
                            xlWorkSheet.Range["Q" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                            xlWorkSheet.Range["Q" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //APP
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_APP), "1");
                            xlWorkSheet.Range["R" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                            xlWorkSheet.Range["R" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Credito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_CREDITO), "1");
                            xlWorkSheet.Range["S" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["S" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Total cajero
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(TOTAL_CAJERO), "1");
                            xlWorkSheet.Range["T" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_CAJERO));
                            xlWorkSheet.Range["T" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message, ex.StackTrace);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                }
                if (Tipo == "3")
                {
                    try
                    {

                        List<RESUMEN_CAJA> ViasPago = new List<RESUMEN_CAJA>();
                        for (int k = 0; k < DGResumenCajasRep.Items.Count; k++)
                        {

                            if (k == 0)
                            {
                                DGResumenCajasRep.Items.MoveCurrentToFirst();
                            }
                            if (DGResumenCajasRep.Items.CurrentItem != null)
                            {
                                ViasPago.Add(DGResumenCajasRep.Items.CurrentItem as RESUMEN_CAJA);
                            }
                            DGResumenCajasRep.Items.MoveCurrentToNext();
                        }
                        //NOMBRES DE COLUMNAS
                        xlWorkSheet.Range["A10"].Value = "Id. sucursal";
                        xlWorkSheet.Range["B10"].Value = "Sucursal";
                        xlWorkSheet.Range["C10"].Value = "Id.  caja";
                        xlWorkSheet.Range["D10"].Value = "Caja";
                        xlWorkSheet.Range["E10"].Value = "Efectivo";
                        xlWorkSheet.Range["F10"].Value = "Chq. al día";
                        xlWorkSheet.Range["G10"].Value = "Chq. a fecha";
                        xlWorkSheet.Range["H10"].Value = "Transferencia";
                        xlWorkSheet.Range["I10"].Value = "V. vista";
                        xlWorkSheet.Range["J10"].Value = "Depósitos";
                        xlWorkSheet.Range["K10"].Value = "Tarjetas";
                        xlWorkSheet.Range["L10"].Value = "Financiamiento";
                        xlWorkSheet.Range["M10"].Value = "APP";
                        xlWorkSheet.Range["N10"].Value = "Crédito";


                        int lineabase = 11;
                        int lineatope = lineabase + ViasPago.Count;

                        FormatoMonedas FM = new FormatoMonedas();
                        string MonedaFormateada = "";
                        for (int k = 0; k < ViasPago.Count; k++)
                        {
                            int linea = lineabase + k;
                            xlWorkSheet.Range["A" + Convert.ToString(linea)].Value = ViasPago[k].ID_SUCURSAL;
                            xlWorkSheet.Range["B" + Convert.ToString(linea)].Value = ViasPago[k].SUCURSAL;
                            xlWorkSheet.Range["C" + Convert.ToString(linea)].Value = ViasPago[k].ID_CAJA;
                            xlWorkSheet.Range["D" + Convert.ToString(linea)].Value = ViasPago[k].NOM_CAJA;
                            //Efectivo
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_EFEC, "2");
                                xlWorkSheet.Range["E" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                xlWorkSheet.Range["E" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                            //Cheque al dia
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_DIA, "2");
                                xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                            //Cheque a fecha
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_FECHA, "2");
                                xlWorkSheet.Range["G" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                xlWorkSheet.Range["G" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                            //Transferencias
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_TRANSF, "2");
                                xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                            //Vale vistas
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_VALE_V, "2");
                                xlWorkSheet.Range["I" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                xlWorkSheet.Range["I" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                            //Depositos
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_DEP, "2");
                                xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                            //Tarjetas
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_TARJ, "2");
                                xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                            //Financiamiento
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_FINANC, "2");
                                xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                            //APP
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_APP, "2");
                                xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                            //Credito
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMonedaChilena2(ViasPago[k].MONTO_CREDITO, "2");
                                xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                        }
                        //TOTALES
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope)].Value = "Totales: ";
                        //Efectivo
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_EFEC), "1");
                            xlWorkSheet.Range["E" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                            xlWorkSheet.Range["E" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque al dia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_DIA), "1");
                            xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                            xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque a fecha
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_FECHA), "1");
                            xlWorkSheet.Range["G" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                            xlWorkSheet.Range["G" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Transferencia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_TRANSF), "1");
                            xlWorkSheet.Range["H" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                            xlWorkSheet.Range["H" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Vale vista
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_VALE_V), "1");
                            xlWorkSheet.Range["I" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                            xlWorkSheet.Range["I" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Deposito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_DEP), "1");
                            xlWorkSheet.Range["J" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                            xlWorkSheet.Range["J" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Tarjetas
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_TARJ), "1");
                            xlWorkSheet.Range["K" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                            xlWorkSheet.Range["K" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Financiamiento
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_FINANC), "1");
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //APP
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_APP), "1");
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Credito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMonedaChilena2(Convert.ToString(MONTO_CREDITO), "1");
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message, ex.StackTrace);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                }
                if (Tipo == "1")
                {
                    xlApp.Visible = true;
                }
                if (Tipo == "2")
                {
                    xlApp.Visible = true;
                }
                if (Tipo == "3")
                {
                    xlApp.Visible = true;
                }

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                //System.Windows.Forms.MessageBox.Show("Archivo Excel creado");
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        private void ImpresionReporteCaja(List<RENDICION_CAJA> ListRendicionCaja, List<RESUMEN_MENSUAL> ListResumenMensual, List<RESUMEN_CAJA> ListResumenCaja, string SociedadR, string Empresa, string Sucursal
          , string RUT, string FechaDesde, string FechaHasta, string Tipo)
        {
            try
            {
                string fecha = Convert.ToString(DateTime.Now);
                fecha = fecha.Replace(" ", "-");
                //fecha = fecha.Replace("-", "");              
                fecha = fecha.Replace(":", "-");
                //fecha = fecha.Substring(0, 19);
                //SE HACE UN PRIMER ARCHIVO QUE SE ALMACENA EN LA VARIABLE direct, ESTE ARCHIVO SIRVE PARA DETERMINAR EL TAMAÑO DEFINITIVO DEL PDF A 
                //REALIZAR, POSTERIORMENTE SE LEE ESTE ARCHIVO PRELIMINAR Y SE CREA UNO NUEVO DONDE EL ENCABEZADO Y EL PIE DE PAGINA SE ESTAMPAN COMO 
                //COMO UNA ESPECIE DE SELLO EN TODAS LAS PAGINAS DEL PDF, ESTO PERMITE INCORPORAR EL NUMERO DE PAGINAS TOTALES, LA PAGINA ACTUAL Y LOS DEMAS 
                //DATOS QUE DESEAN REPETIRSE
                string appRootDir = Convert.ToString(System.IO.Path.GetTempPath());
                string watermarkedFile = "";
                string direct = Convert.ToString(System.IO.Path.GetTempPath());

                if (Tipo == "1")
                {
                    direct = direct + "inchcapeLog\\" + "RendicionCaja" + fecha + ".pdf";
                    watermarkedFile = appRootDir + "inchcapeLog\\" + "RendicionCaja" + fecha + "-Nuevo.Text.pdf";

                }
                if (Tipo == "2")
                {
                    direct = direct + "inchcapeLog\\" + "ResumenMensualMovimientos" + fecha + ".pdf";
                    watermarkedFile = appRootDir + "inchcapeLog\\" + "ResumenMensualMovimientos" + fecha + "-Nuevo.Text.pdf";
                }
                if (Tipo == "3")
                {
                    direct = direct + "inchcapeLog\\" + "ResumenCaja" + fecha + ".pdf";
                    watermarkedFile = appRootDir + "inchcapeLog\\" + "ResumenCaja" + fecha + "-Nuevo.Text.pdf";
                }
                //MARCA DE AGUA PARA DOCUMENTOS DE PRUEBA O SIN VALIDEZ
                string watermarkText = "Documento No válido";
                string Cajero = "";

                using (FileStream fs = new FileStream(direct, FileMode.Create, FileAccess.Write, FileShare.None))

                using (Document pdfcommande = new Document(PageSize.LETTER.Rotate(), 20f, 20f, 100f, 100f))

                using (PdfWriter writer = PdfWriter.GetInstance(pdfcommande, fs))
                {
                    txtDirect.Text = direct;
                    try
                    {
                        pdfcommande.Open();

                        pdfcommande.NewPage();
                        string Titulo = "";
                        if (Tipo == "1")
                        {
                            Titulo = "Informe Rendición de caja";
                        }
                        if (Tipo == "2")
                        {
                            Titulo = "Informe Resumen de mensual de movimientos";
                        }
                        if (Tipo == "3")
                        {
                            Titulo = "Resumen Caja Recaudadora";
                        }


                        //Titulo
                        string texto = Titulo;
                        iTextSharp.text.Paragraph itxtTitulo = new iTextSharp.text.Paragraph(texto);
                        itxtTitulo.IndentationLeft = 300;
                        itxtTitulo.Font.Size = 10;
                        itxtTitulo.Font.SetStyle(1);
                        itxtTitulo.Font.SetFamily("Courier");
                        itxtTitulo.SpacingBefore = 30f;
                        itxtTitulo.SpacingAfter = 5f;
                        pdfcommande.Add(itxtTitulo);
                        //Fecha desde hasta del Informe
                        //if (Tipo != "1")
                        //{
                        texto = "Desde: " + FechaDesde + "           " + "Hasta: " + FechaHasta;
                        iTextSharp.text.Paragraph itxtfechDesdeHasta = new iTextSharp.text.Paragraph(texto);
                        itxtfechDesdeHasta.IndentationLeft = 270;
                        itxtfechDesdeHasta.Font.Size = 9;
                        itxtfechDesdeHasta.Font.SetFamily("Courier");
                        itxtfechDesdeHasta.SpacingAfter = 5f;
                        pdfcommande.Add(itxtfechDesdeHasta);
                        //}

                        //Datos Caja
                        texto = "Id Caja:" + Convert.ToString(textBlock6.Content);
                        iTextSharp.text.Paragraph itxtIngreso = new iTextSharp.text.Paragraph(texto);
                        itxtIngreso.IndentationLeft = 15;
                        itxtIngreso.Font.Size = 9;
                        itxtIngreso.Font.SetFamily("Courier");
                        pdfcommande.Add(itxtIngreso);
                        //Datos de nota venta
                        texto = "Sucursal" + Convert.ToString(textBlock8.Content);
                        iTextSharp.text.Paragraph itxtNotaVta = new iTextSharp.text.Paragraph(texto);
                        itxtNotaVta.IndentationLeft = 15;
                        itxtNotaVta.Font.Size = 9;
                        itxtNotaVta.Font.SetFamily("Courier");
                        itxtNotaVta.SpacingAfter = 10;
                        pdfcommande.Add(itxtNotaVta);


                        PdfPTable table2;

                        double MONTO = 0;
                        double TOTAL_MOV = 0;
                        double TOTAL_INGR = 0;
                        double MONTO_EFEC = 0;
                        double MONTO_DIA = 0;
                        double MONTO_FECHA = 0;
                        double MONTO_TRANSF = 0;
                        double MONTO_VALE_V = 0;
                        double MONTO_DEP = 0;
                        double MONTO_TARJ = 0;
                        double MONTO_FINANC = 0;
                        double MONTO_APP = 0;
                        double MONTO_CREDITO = 0;
                        double TOTAL_CAJERO = 0;
                        FormatoMonedas FM = new FormatoMonedas();
                        string MonedaFormateada;
                        if (Tipo == "1")
                        {
                            DGRendicionCajaRep.ItemsSource = null;
                            DGRendicionCajaRep.Items.Clear();
                            DGRendicionCajaRep.ItemsSource = ListRendicionCaja;

                            if (ListRendicionCaja.Count > 0)
                            {
                                //DEFINE CUANTAS COLUMNAS LLEVA LA GRILLA
                                table2 = new PdfPTable(DGRendicionCajaRep.Columns.Count);
                                table2.TotalWidth = 775f;
                                string tableheight = Convert.ToString(table2.TotalHeight);
                                table2.LockedWidth = true;
                                table2.HeaderRows = 2;
                                table2.SpacingAfter = 30f;
                                //ANCHO DE LAS COLUMNAS
                                float[] widths = new float[] { 20f, 50f, 0f, 50f, 50f, 50f, 120f, 50f, 50f, 50f, 50f, 50f, 50f, 50f, 50f, 50f, 50f, 50f, 50f };
                                table2.SetWidths(widths);
                                //LISTA CON LOS DATOS QUE SE MOSTRARAN EN  LA GRILLA
                                List<RENDICION_CAJA> ViasPago = new List<RENDICION_CAJA>();
                                for (int k = 0; k < DGRendicionCajaRep.Items.Count; k++)
                                {
                                    if (k == 0)
                                    {
                                        DGRendicionCajaRep.Items.MoveCurrentToFirst();
                                    }
                                    ViasPago.Add(DGRendicionCajaRep.Items.CurrentItem as RENDICION_CAJA);
                                    DGRendicionCajaRep.Items.MoveCurrentToNext();
                                }
                                Cajero = ViasPago[0].CAJERO;
                                PdfPCell cell2 = new PdfPCell(new Phrase(Convert.ToString(label11.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                                cell2.Padding = 10f;
                                cell2.Colspan = DGRendicionCajaRep.Columns.Count;
                                cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                table2.AddCell(cell2);
                                //DEFINICION DE ENCABEZADO DE LA GRILLA
                                foreach (DataGridColumn column in DGRendicionCajaRep.Columns)
                                {
                                    table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));

                                }
                                //LLENADO DE LA GRILLA
                                for (int k = 0; k <= ViasPago.Count - 1; k++)
                                {
                                    //if (ViasPago[k] != null)
                                    //{
                                    try
                                    {
                                        //HASTA LA POSICION MAXIMA DE LA LISTA - 1, SE MUESTRAN LOS DATOS DE LA LISTA.
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].N_VENTA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        //EN LA POSICION MAXIMA, O ES VACIO, O SE COLOCA ALGUN TITULO O SE MUESTRAN LOS TOTALES DE LAS COLUMNAS CON MONTOS.
                                        else
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow20 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DOC_TRIB), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow20);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow20 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow20);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow21 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].CAJERO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow21);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow21 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow21);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FEC_EMI), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FEC_VENC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString("Totales"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        //MONTO
                                        //HASTA LA POSICION MAXIMA DE LA LISTA - 1, SE MUESTRAN LOS DATOS DE LA LISTA.
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO, "2");
                                                PdfPCell cellrow4 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow4.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow4);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO);
                                                PdfPCell cellrow4 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow4.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow4);
                                            }
                                            MONTO = MONTO + Convert.ToDouble(ViasPago[k].MONTO);
                                        }
                                        //EN LA POSICION MAXIMA, O ES VACIO, O SE COLOCA ALGUN TITULO O SE MUESTRAN LOS TOTALES DE LAS COLUMNAS CON MONTOS.
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO), "1");
                                                PdfPCell cellrow4 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow4.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow4);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO));
                                                PdfPCell cellrow4 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow4.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow4);
                                            }
                                        }
                                        //NAME1
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NAME1), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow5);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow5);
                                        }
                                        //MONTO_EFEC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_EFEC, "2");
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);

                                                MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }        
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_EFEC), "1");
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }
                                        }
                                        //NUM_CHEQUE
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NUM_CHEQUE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow7);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow7);
                                        }
                                        //MONTO_DIA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_DIA, "2");
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_DIA), "1");
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                        }
                                        //MONTO_FECHA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_FECHA, "2");
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                            MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_FECHA), "1");
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                        }
                                        //MONTO_TRANSF
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_TRANSF, "2");
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_TRANSF), "1");
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                        }
                                        //MONTO_VALE_V
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_VALE_V, "2");
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_VALE_V), "1");
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                        }
                                        //MONTO_DEP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_DEP, "2");
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_DEP), "1");
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                        }
                                        //MONTO_TARJ
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_TARJ, "2");
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_TARJ), "1");
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                        }
                                        //MONTO_FINANC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_FINANC, "2");
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_FINANC), "1");
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                        }
                                        //MONTO_APP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_APP, "2");
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_APP), "1");
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                        }
                                        //MONTO_CREDITO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_CREDITO, "2");
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_CREDITO), "1");
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                        }
                                        //DOC_SAP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow19 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DOC_SAP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow19);

                                        }
                                        else
                                        {
                                            PdfPCell cellrow19 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow19);
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.Write(ex.Message, ex.StackTrace);
                                    }
                                    //}
                                }
                                pdfcommande.Add(table2);
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("No Existen datos Para el intervalo seleccionado");
                            }

                        }

                        //INFORME RESUMEN DE MOVIMIENTOS
                        //EL TRATAMIENTO ES EL MISMO QUE EL INFORME DE RENDICION
                        if (Tipo == "2")
                        {
                            DGResumenMovimientosRep.ItemsSource = null;
                            DGResumenMovimientosRep.Items.Clear();
                            DGResumenMovimientosRep.ItemsSource = ListResumenMensual;

                            if (ListResumenMensual.Count > 0)
                            {

                                table2 = new PdfPTable(DGResumenMovimientosRep.Columns.Count);
                                table2.TotalWidth = 780f;
                                table2.LockedWidth = true;
                                table2.HeaderRows = 2;
                                table2.SpacingBefore = 20f;
                                table2.SpacingAfter = 20f;
                                float[] widths = new float[] { 30f, 100f, 30f, 100f, 40f, 30f, 80f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f };
                                table2.SetWidths(widths);
                                //int factor = 1;
                                PdfPCell cell2 = new PdfPCell(new Phrase(Convert.ToString(label11.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                                cell2.Padding = 10f;
                                cell2.Colspan = DGResumenMovimientosRep.Columns.Count;
                                cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                table2.AddCell(cell2);
                                foreach (DataGridColumn column in DGResumenMovimientosRep.Columns)
                                {
                                    table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));

                                }

                                List<RESUMEN_MENSUAL> ViasPago = new List<RESUMEN_MENSUAL>();
                                for (int k = 0; k < DGResumenMovimientosRep.Items.Count; k++)
                                {

                                    if (k == 0)
                                    {
                                        DGResumenMovimientosRep.Items.MoveCurrentToFirst();
                                    }
                                    ViasPago.Add(DGResumenMovimientosRep.Items.CurrentItem as RESUMEN_MENSUAL);
                                    DGResumenMovimientosRep.Items.MoveCurrentToNext();
                                }


                                for (int k = 0; k < ViasPago.Count; k++)
                                {
                                    try
                                    {
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].ID_SUCURSAL), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        //SUCURSAL
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].SUCURSAL), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        //ID_CAJA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].ID_CAJA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        //NOM_CAJA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NOM_CAJA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow4);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow4);
                                        }
                                        //CAJERO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].CAJERO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow5);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow5);
                                        }
                                        //FLIJO_DOCS
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow6 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FLUJO_DOCS), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow6);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow6 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow6);
                                        }
                                        //AREA_VTAS
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].AREA_VTAS), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow7);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString("Totales"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow7);
                                        }
                                        //TOTAL_MOV
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].TOTAL_MOV, "2");
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_MOV);
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            TOTAL_MOV = TOTAL_MOV + Convert.ToDouble(ViasPago[k].TOTAL_MOV);
                                        }
                                        else
                                        {
                                            //PdfPCell cellrow8 = new PdfPCell(new Phrase(Convert.ToString(TOTAL_MOV), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            //cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            //table2.AddCell(cellrow8);
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(TOTAL_MOV), "1");
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_MOV));
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                        }
                                        //TOTAL_INGR
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].TOTAL_INGR, "2");
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_INGR);
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                            TOTAL_INGR = TOTAL_INGR + Convert.ToDouble(ViasPago[k].TOTAL_INGR);
                                        }
                                        else
                                        {
                                            //PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(TOTAL_INGR), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                            //cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            //table2.AddCell(cellrow9);
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(TOTAL_INGR), "1");
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_INGR));
                                                PdfPCell cellrow9 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow9);
                                            }
                                        }
                                        //MONTO_EFEC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_EFEC, "2");
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_EFEC), "1");
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                        }
                                        //MONTO_DIA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_DIA, "2");
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_DIA), "1");
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                        }
                                        //MONTO_FECHA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_FECHA, "2");
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_FECHA), "1");
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                        }
                                        //MONTO_TRANSF
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_TRANSF, "2");
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_TRANSF), "1");
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                        }
                                        //MONTO_VALE_V
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_VALE_V, "2");
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_VALE_V), "1");
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                        }
                                        //MONTO_DEP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_DEP, "2");
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_DEP), "1");
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                        }
                                        //MONTO_TARJ
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_TARJ, "2");
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_TARJ), "1");
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                        }
                                        //MONTO_FINANC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_FINANC, "2");
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_FINANC), "1");
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                        }
                                        //MONTO_APP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_APP, "2");
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_APP), "1");
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                        }
                                        //MONTO_CREDITO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_CREDITO, "2");
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_CREDITO), "1");
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                        }
                                        //TOTAL_CAJERO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].TOTAL_CAJERO, "2");
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);

                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_CAJERO);
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                            TOTAL_CAJERO = TOTAL_CAJERO + Convert.ToDouble(ViasPago[k].TOTAL_CAJERO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(TOTAL_CAJERO), "1");
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);

                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_CAJERO));
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.Write(ex.Message, ex.StackTrace);
                                    }
                                }
                                pdfcommande.Add(table2);
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("No Existen datos Para el intervalo seleccionado");
                            }
                        }
                        if (Tipo == "3")
                        {
                            DGResumenCajasRep.ItemsSource = null;
                            DGResumenCajasRep.Items.Clear();
                            DGResumenCajasRep.ItemsSource = ListResumenCaja;

                            if (ListResumenCaja.Count > 0)
                            {

                                table2 = new PdfPTable(DGResumenCajasRep.Columns.Count);
                                table2.TotalWidth = 775f;
                                table2.LockedWidth = true;
                                table2.HeaderRows = 2;
                                table2.SpacingAfter = 30f;
                                float[] widths = new float[] { 30f, 100f, 30f, 100f, 60f, 60f, 60f, 60f, 60f, 60f, 60f, 60f, 60f, 60f };
                                table2.SetWidths(widths);
                                List<RESUMEN_CAJA> ViasPago = new List<RESUMEN_CAJA>();
                                for (int k = 0; k < DGResumenCajasRep.Items.Count; k++)
                                {

                                    if (k == 0)
                                    {
                                        DGResumenCajasRep.Items.MoveCurrentToFirst();
                                    }
                                    ViasPago.Add(DGResumenCajasRep.Items.CurrentItem as RESUMEN_CAJA);
                                    DGResumenCajasRep.Items.MoveCurrentToNext();
                                }

                                PdfPCell cell2 = new PdfPCell(new Phrase(Convert.ToString(label11.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                                cell2.Padding = 10f;
                                cell2.Colspan = DGResumenCajasRep.Columns.Count;
                                cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                table2.AddCell(cell2);
                                foreach (DataGridColumn column in DGResumenCajasRep.Columns)
                                {
                                    table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));

                                }

                                for (int k = 0; k < ViasPago.Count; k++)
                                {
                                    try
                                    {
                                        //ID_SUCURSAL
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].ID_SUCURSAL), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        //SUCURSAL
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].SUCURSAL), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        //ID_CAJA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].ID_CAJA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        //NOM_CAJA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NOM_CAJA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow4);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString("Totales"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow4);
                                        }
                                        //MONTO_EFEC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_EFEC, "2");
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_EFEC), "1");
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                        }
                                        //MONTO_DIA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_DIA, "2");
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_DIA), "1");
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                        }
                                        //MONTO_FECHA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_FECHA, "2");
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_FECHA), "1");
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                        }
                                        //MONTO_TRANSF
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_TRANSF, "2");
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_TRANSF), "1");
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                        }
                                        //MONTO_VALE_V
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_VALE_V, "2");
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_VALE_V), "1");
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                        }
                                        //MONTO_DEP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_DEP, "2");
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_DEP), "1");
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                        }
                                        //MONTO_TARJ
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_TARJ, "2");
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_TARJ), "1");
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                        }
                                        //MONTO_FINANC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_FINANC, "2");
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_FINANC), "1");
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                        }
                                        //MONTO_APP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_APP, "2");
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_APP), "1");
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                        }
                                        //MONTO_CREDITO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_CREDITO, "2");
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMonedaChilena(Convert.ToString(MONTO_CREDITO), "1");
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                        }

                                    }
                                    catch (Exception ex)
                                    {
                                        Console.Write(ex.Message, ex.StackTrace);
                                    }
                                }
                                pdfcommande.Add(table2);
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("No Existen datos Para el intervalo seleccionado");
                            }
                        }
                        pdfcommande.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message, ex.StackTrace);
                    }
                }
                try
                {
                    string direct2 = Convert.ToString(System.IO.Path.GetTempPath());

                    PdfReader reader1 = new PdfReader(direct);
                    using (FileStream fs = new FileStream(watermarkedFile, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (PdfStamper stamper = new PdfStamper(reader1, fs))
                    {
                        int pageCount = reader1.NumberOfPages;
                        PdfLayer layer = new PdfLayer("WatermarkLayer", stamper.Writer);
                        for (int i = 1; i <= pageCount; i++)
                        {
                            string PaginaActual = Convert.ToString(i);
                            string PaginasTotales = Convert.ToString(pageCount);
                            iTextSharp.text.Rectangle rect = reader1.GetPageSize(i);
                            rect.Rotate();
                            if (txtMandante.Text != "300")
                            {
                                PdfContentByte cb = stamper.GetUnderContent(i);
                                cb.BeginLayer(layer);
                                cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 50);
                                PdfGState gState = new PdfGState();
                                gState.FillOpacity = 0.25f;
                                cb.SetGState(gState);
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, watermarkText, rect.Width / 2, rect.Height / 2, 45f);
                                cb.EndText();
                                cb.EndLayer();
                            }

                            //PAGINACION
                            PdfContentByte cb2 = stamper.GetUnderContent(i);
                            cb2.BeginLayer(layer);
                            cb2.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            PdfGState gState2 = new PdfGState();
                            gState2.FillOpacity = 1f;
                            cb2.SetGState(gState2);
                            cb2.SetColorFill(BaseColor.BLACK);
                            cb2.BeginText();
                            string Paginas = "Pagina: " + PaginaActual + "         De: " + PaginasTotales;
                            cb2.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Paginas, rect.Width, rect.Height - 200, 0f);
                            cb2.EndText();
                            cb2.EndLayer();

                            //EMPRESA
                            PdfContentByte cb3 = stamper.GetUnderContent(i);
                            cb3.BeginLayer(layer);
                            cb3.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb3.SetGState(gState2);
                            cb3.SetColorFill(BaseColor.BLACK);
                            cb3.BeginText();
                            cb3.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Empresa, rect.Width - (rect.Width - 10), rect.Height - 200, 0f);
                            cb3.EndText();
                            cb3.EndLayer();

                            //FECHA LABEL
                            PdfContentByte cb4a = stamper.GetUnderContent(i);
                            cb4a.BeginLayer(layer);
                            cb4a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb4a.SetGState(gState2);
                            cb4a.SetColorFill(BaseColor.BLACK);
                            cb4a.BeginText();
                            cb4a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Fecha:", rect.Width, rect.Height - 210, 0f);
                            cb4a.EndText();
                            cb4a.EndLayer();
                            //FECHA
                            PdfContentByte cb4 = stamper.GetUnderContent(i);
                            cb4.BeginLayer(layer);
                            cb4.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb4.SetGState(gState2);
                            cb4.SetColorFill(BaseColor.BLACK);
                            cb4.BeginText();
                            string FechaAFormatear = Convert.ToString(DateTime.Now).Substring(0, 10);
                            cb4.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, String.Format("{0:dd/MM/yyyy}", FechaAFormatear), rect.Width + 160, rect.Height - 210, 0f);
                            cb4.EndText();
                            // Close the layer
                            cb4.EndLayer();

                            // RUT EMPRESA
                            PdfContentByte cb5 = stamper.GetUnderContent(i);
                            cb5.BeginLayer(layer);
                            cb5.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb5.SetGState(gState2);
                            cb5.SetColorFill(BaseColor.BLACK);
                            cb5.BeginText();
                            cb5.ShowTextAligned(PdfContentByte.ALIGN_LEFT, RUT, rect.Width - (rect.Width - 10), rect.Height - 210, 0f);
                            cb5.EndText();
                            cb5.EndLayer();

                            //HORA
                            PdfContentByte cb6 = stamper.GetUnderContent(i);
                            cb6.BeginLayer(layer);
                            cb6.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb6.SetGState(gState2);
                            cb6.SetColorFill(BaseColor.BLACK);
                            cb6.BeginText();
                            FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                            cb6.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, String.Format("{0:HH:mm:ss}", FechaAFormatear), rect.Width + 160, rect.Height - 220, 0f);
                            cb6.EndText();
                            cb6.EndLayer();

                            //HORA LABEL
                            PdfContentByte cb6a = stamper.GetUnderContent(i);
                            cb6a.BeginLayer(layer);
                            cb6a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb6a.SetGState(gState2);
                            cb6a.SetColorFill(BaseColor.BLACK);
                            cb6a.BeginText();
                            cb6a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Hora:", rect.Width, rect.Height - 220, 0f);
                            cb6a.EndText();
                            cb6a.EndLayer();

                            // SOCIEDAD
                            PdfContentByte cb7 = stamper.GetUnderContent(i);
                            cb7.BeginLayer(layer);
                            cb7.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb7.SetGState(gState2);
                            cb7.SetColorFill(BaseColor.BLACK);
                            cb7.BeginText();
                            cb7.ShowTextAligned(PdfContentByte.ALIGN_LEFT, SociedadR, rect.Width - (rect.Width - 10), rect.Height - 220, 0f);
                            cb7.EndText();
                            cb7.EndLayer();

                            //USUARIO
                            PdfContentByte cb8 = stamper.GetUnderContent(i);
                            cb8.BeginLayer(layer);
                            cb8.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb8.SetGState(gState2);
                            cb8.SetColorFill(BaseColor.BLACK);
                            cb8.BeginText();
                            FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                            cb8.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, Convert.ToString(textBlock7.Content), rect.Width + 160, rect.Height - 230, 0f);
                            cb8.EndText();
                            cb8.EndLayer();

                            //USUARIO LABEL
                            PdfContentByte cb8a = stamper.GetUnderContent(i);
                            cb8a.BeginLayer(layer);
                            cb8a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                            gState2.FillOpacity = 1f;
                            cb8a.SetGState(gState2);
                            cb8a.SetColorFill(BaseColor.BLACK);
                            cb8a.BeginText();
                            cb8a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Usuario:", rect.Width, rect.Height - 230, 0f);
                            cb8a.EndText();
                            cb8a.EndLayer();

                            if (Tipo == "1")
                            {
                                //CAJERO
                                PdfContentByte cb9 = stamper.GetUnderContent(i);
                                cb9.BeginLayer(layer);
                                cb9.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb9.SetGState(gState2);
                                cb9.SetColorFill(BaseColor.BLACK);
                                cb9.BeginText();
                                cb9.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, Cajero, rect.Width + 160, rect.Height - 240, 0f);
                                cb9.EndText();
                                // Close the layer
                                cb9.EndLayer();

                                //CAJERO LABEL
                                PdfContentByte cb9a = stamper.GetUnderContent(i);
                                cb9a.BeginLayer(layer);
                                cb9a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb9a.SetGState(gState2);
                                cb9a.SetColorFill(BaseColor.BLACK);
                                cb9a.BeginText();
                                cb9a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Cajero:", rect.Width, rect.Height - 240, 0f);
                                cb9a.EndText();
                                // Close the layer
                                cb9a.EndLayer();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.Write(ex.Message, ex.StackTrace);
                }

                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.FileName = watermarkedFile;
                proc.Start();
                proc.Close();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
            }
        }

        public List<DETALLE_ARQUEO> ListaEfectivo()
        {

            List<DETALLE_ARQUEO> DetalleEfectivo = new List<DETALLE_ARQUEO>();
            DETALLE_ARQUEO DetArqueo = new DETALLE_ARQUEO();

            //20000
            if (txtC20000.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "20000";
                DetArqueo.CANTIDAD_MON = txtC20000.Text;
                DetArqueo.SUMA_MON_BILL = txtT20000.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //10000
            if (txtC10000.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "10000";
                DetArqueo.CANTIDAD_MON = txtC10000.Text;
                DetArqueo.SUMA_MON_BILL = txtT10000.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //5000
            if (txtC5000.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "5000";
                DetArqueo.CANTIDAD_MON = txtC5000.Text;
                DetArqueo.SUMA_MON_BILL = txtT5000.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //2000
            if (txtC2000.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "2000";
                DetArqueo.CANTIDAD_MON = txtC2000.Text;
                DetArqueo.SUMA_MON_BILL = txtT2000.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //1000
            if (txtC1000.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "1000";
                DetArqueo.CANTIDAD_MON = txtC1000.Text;
                DetArqueo.SUMA_MON_BILL = txtT1000.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //500
            if (txtC500.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "500";
                DetArqueo.CANTIDAD_MON = txtC500.Text;
                DetArqueo.SUMA_MON_BILL = txtT500.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //100
            if (txtC100.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "100";
                DetArqueo.CANTIDAD_MON = txtC100.Text;
                DetArqueo.SUMA_MON_BILL = txtT100.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //50
            if (txtC50.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "50";
                DetArqueo.CANTIDAD_MON = txtC50.Text;
                DetArqueo.SUMA_MON_BILL = txtT50.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //10
            if (txtC10.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "10";
                DetArqueo.CANTIDAD_MON = txtC10.Text;
                DetArqueo.SUMA_MON_BILL = txtT10.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //5
            if (txtC5.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "5";
                DetArqueo.CANTIDAD_MON = txtC5.Text;
                DetArqueo.SUMA_MON_BILL = txtT5.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            DetArqueo = new DETALLE_ARQUEO();
            //1
            if (txtC1.Text != "0")
            {
                DetArqueo.LAND = Convert.ToString(lblPais.Content);
                DetArqueo.ID_CAJA = Convert.ToString(textBlock6.Content);
                DetArqueo.USUARIO = Convert.ToString(textBlock7.Content);
                DetArqueo.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                DetArqueo.HORA_REND = Convert.ToString(DateTime.Now);
                DetArqueo.FECHA_REND = Convert.ToString(calendario.SelectedDate);
                DetArqueo.MONEDA = logApertura2[0].MONEDA;
                DetArqueo.VIA_PAGO = "E";
                DetArqueo.TIPO_MONEDA = "1";
                DetArqueo.CANTIDAD_MON = txtC1.Text;
                DetArqueo.SUMA_MON_BILL = txtT1.Text;
                DetArqueo.CANTIDAD_DOC = "0000000000";
                DetArqueo.SUMA_DOCS = "0,00";
                DetalleEfectivo.Add(DetArqueo);
            }
            return DetalleEfectivo;
            GC.Collect();
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                System.Windows.Forms.MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void PagoDocumentos_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void EmisionNC_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void Anulacion_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void Reimpresion_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void RecaudacionVeh_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void CierreCajas_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }
        private void BloquearCaja_Click(object sender, RoutedEventArgs e)
        {
            bloquearCaja();
        }

        private void bloquearCaja()
        {
            timer.Stop();
            BloquearCaja bloquearcaja = new BloquearCaja();
            bloquearcaja.bloqueardesbloquearcaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, logApertura2);
            this.Close();
            Login = new MainWindow();
            Login.Show();
            GC.Collect();
        }

        private void ReportesCaja_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }
    }
}
