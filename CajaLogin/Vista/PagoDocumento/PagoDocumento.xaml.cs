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
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections.ObjectModel;
using System.Collections;
using System.Windows.Controls.Primitives;
using iTextSharp.text;
using iTextSharp.text.pdf;
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

namespace CajaIndigo.Vista.PagoDocumento
{
    /// <summary>
    /// Interaction logic for PagoDocumento.xaml
    /// </summary>
    public partial class PagoDocumento : System.Windows.Window
    {
        public PagoDocumento()
        {
            InitializeComponent();
        }

        private PdfTemplate totalPages;
        private PdfWriter Write;
        List<DetalleViasPago> cheques = new List<DetalleViasPago>();
        List<VIAS_PAGO_MASIVO> chequesMasiv = new List<VIAS_PAGO_MASIVO>();
        List<T_DOCUMENTOS> detalledocs = new List<T_DOCUMENTOS>();
        List<T_DOCUMENTOS> partidaseleccionadas = new List<T_DOCUMENTOS>();
        List<T_DOCUMENTOSAUX> partidaseleccionadasaux = new List<T_DOCUMENTOSAUX>();
        public List<ViasPago> ViasPagoTransaccion = new List<ViasPago>();
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
        List<LOG_APERTURA> logApertura = new List<LOG_APERTURA>();
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();
        List<VALIDAREFECTIVO> ValidEfec = new List<VALIDAREFECTIVO>();
        int suma = 0;
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
        double monto;
        double monto2;
        double monto3;
        string Valor2 = string.Empty;
        string moneda = string.Empty;

        public string Term { get; set; }

        Vista.PagoDocumento.PagoDocumento PagDocum;
        Vista.NotaCredito.NotaCredito NotaCredit;
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.Reportes.Reportes Reporte;

        FormatoMonedas Formato = new FormatoMonedas();

        public PagoDocumento(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
        {
            try
            {
                InitializeComponent();
                List<string> myItemsCollection = new List<string>();
                myItemsCollection.Add(moneda);
                int test = 0;
                test = cmbMoneda.Items.Count;
                GBInicio.Visibility = Visibility.Collapsed;
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
                cmbMoneda.Items.Clear();
                cmbMonedaMasivo.Items.Clear();
                cmbMoneda.ItemsSource = myItemsCollection;
                cmbMonedaMasivo.ItemsSource = myItemsCollection;
                if (cmbMoneda.SelectedValue != "0" && cmbMoneda.SelectedValue != "0")
                {
                    cmbMoneda.SelectedIndex = 0;
                    cmbMonedaMasivo.SelectedIndex = 0;
                }
                lblPais.Content = pais;
                lblPassword.Content = passlogg;
                DateTime result = DateTime.Today;
                Calendario.Text = Convert.ToString(result);
                DGLogApertura.ItemsSource = null;
                DGLogApertura.Items.Clear();
                DGLogApertura.ItemsSource = logApertura;
                logApertura2 = logApertura;
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
        private void Window_Loaded()
        {
            GBInicio.Visibility = Visibility.Visible;
            GBMonitor.Visibility = Visibility.Visible;
            DateTime result = DateTime.Today;
            Calendario.Text = Convert.ToString(result);
            //ACTIVACION DEL MONITOR
            chkMonitor.IsChecked = true;
            //RFC PARA OBTENER LOS BANCOS
            RFC_Combo_Bancos();
            GC.Collect();
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
            DGLogApertura.ItemsSource = logApertura;

            if (PagoDocumentos.IsMouseOver == true)
            {
                PagDocum = new PagoDocumento(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
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
                Anula = new Anulacion.Anulacion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Anula.Show();
                this.Visibility = Visibility.Collapsed;
            }
            if (Reimpresion.IsMouseOver == true)
            {
                Reimp = new Reimpresion.Reimpresion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Reimp.Show();
                this.Hide(); ;
            }

            if (RecaudacionVeh.IsMouseOver == true)
            {
                Vehi = new Vehiculos.Vehiculo(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Vehi.Show();
                this.Hide();
            }

            if (CierreCaja.IsMouseOver == true)
            {
                CierCaja = new CierreCaja.CierreCaja(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
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

        private void Window_Closed(object sender, EventArgs e)
        {
            timer.Stop();
            MainWindow window = System.Windows.Window.GetWindow(this.Owner) as MainWindow;
            if (window != null)
            {
                this.Close();
                window.Visibility = Visibility.Visible;
            }
        }

        //MANEJO DE LOS EVENTOS ASOCIADOS A TABCONTROLS
        #region TabControl
        private void tabControlAnulacion_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            txtDocuAnt.Text = "";
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            GC.Collect();
            txtDocuAnt.Text = "";
            txtRut.Text = "";
            txtRUTAnt.Text = "";
            txtArchivo.Text = "";
            GC.Collect();
        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            txtDocu.Text = "";
            txtDocuAnt.Text = "";
            txtRut.Text = "";
            txtRUTAnt.Text = "";
            GC.Collect();
        }
        #endregion

        private void ChkDocsPagar_Checked(object sender, RoutedEventArgs e)
        {
            List<T_DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOSAUX>();
            partidaseleccionadasaux2.Clear();
            if (this.DGPagos.Items.Count > 0)
            {
                for (int i = 0; i < DGPagos.Items.Count; i++)
                {
                    if (i == 0)
                        DGPagos.Items.MoveCurrentToFirst();
                    {
                        partidaseleccionadasaux2.Add(DGPagos.Items.CurrentItem as T_DOCUMENTOSAUX);
                    }
                    DGPagos.Items.MoveCurrentToNext();
                }
            }
            for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
            {
                partidaseleccionadasaux2[k].ISSELECTED = true;
            }
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
            DGPagos.ItemsSource = partidaseleccionadasaux2;
            GC.Collect();
        }
        private void ChkDocsPagar_Unchecked(object sender, RoutedEventArgs e)
        {
            List<T_DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOSAUX>();
            partidaseleccionadasaux2.Clear();
            if (this.DGPagos.Items.Count > 0)
            {
                for (int i = 0; i < DGPagos.Items.Count; i++)
                {
                    if (i == 0)
                        DGPagos.Items.MoveCurrentToFirst();
                    {
                        partidaseleccionadasaux2.Add(DGPagos.Items.CurrentItem as T_DOCUMENTOSAUX);
                    }
                    DGPagos.Items.MoveCurrentToNext();
                }
            }
            for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
            {
                partidaseleccionadasaux2[k].ISSELECTED = false;
            }
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
            DGPagos.ItemsSource = partidaseleccionadasaux2;
            GC.Collect();
        }
        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {
            GC.Collect();
        }

        //MANEJO DE LOS EVENTOS ASOCIADOS A LOS BOTONES
        #region Botones

        //CONEXION A LA RFC DEL MONITOR EN MODO MANUAL
        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //RFC del Monitor por boton Refresh 
                monitor.ObjDatosMonitor.Clear();
                monitor.monitor(Convert.ToString(Calendario.SelectedDate.Value), Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblSociedad.Content));
                if (monitor.ObjDatosMonitor.Count > 0)
                {
                    DGMonitor.ItemsSource = null;
                    DGMonitor.Items.Clear();
                    DGMonitor.ItemsSource = monitor.ObjDatosMonitor;
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
            }
        }

        //CLICK BOTON DE BARRA DE HERRAMIENTAS QUE ACTIVA EL PAGO DE DOCUMENTOS
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Visible;
            GBInicio.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;

            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            GC.Collect();
        }
        //CLICK BOTON DE BARRA DE HERRAMIENTAS QUE ACTIVA LA ANULACION DE DOCUMENTOS
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            CargarDatos();
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            GC.Collect();
        }


        //CLICK BOTON DE MENU QUE ACTIVA LA REIMPRESION DE DOCUMENTOS
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            CargarDatos();
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            btnAutAnul.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            GC.Collect();
        }
        //BOTON DE MENU PARA EL CIERRE DE CAJA
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            CargarDatos();
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBMonitor.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarElementosDeCierreDeCaja();
            //LimpiarCamposInformeRendicion();
            LimpiarViasDePago();
            LimpiarEntradasDeDatos();
        }
        //MENU RECAUDACION DE VEHICULO
        private void bt_recaudacion(object sender, RoutedEventArgs e)
        {
            CargarDatos();
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarEntradasDeDatos();
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            GC.Collect();
        }
        //MENU GESTION DE DEPOSITOS
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarEntradasDeDatos();
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            GC.Collect();
        }
        private void btnPagosMasivo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LimpiarPagMasivo();// Limpia Cajas de Texto
                cmbVPMedioPagMasivo.ItemsSource = null;
                cmbVPMedioPagMasivo.Items.Clear();
                if (txtArchivo.Text != "")
                {
                    GBViasPagoMasivos.Visibility = Visibility.Collapsed;
                    List<ViasPago> Condiciones = new List<ViasPago>();
                    List<string> CondicionPago = new List<string>();

                    ViasPago Condic;// = new ViasPago(acc, cond_pago, caja);
                    Condic = new ViasPago("", "TODO", Convert.ToString(textBlock6.Content));
                    Condiciones.Add(Condic);

                    MatrizDePago matrizpago = new MatrizDePago();
                    matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "", "D", "", Convert.ToString(lblPais.Content), "", Condiciones);

                    if (matrizpago.ObjDatosViasPago.Count > 0)
                    {
                        List<string> VP = new List<string>();
                        for (int i = 0; i < matrizpago.ObjDatosViasPago.Count; i++)
                        {
                            if (VP.Contains(matrizpago.ObjDatosViasPago[i].VIA_PAGO + " - " + matrizpago.ObjDatosViasPago[i].DESCRIPCION))
                            {
                                ;
                            }
                            else
                            {
                                VP.Add(matrizpago.ObjDatosViasPago[i].VIA_PAGO + " - " + matrizpago.ObjDatosViasPago[i].DESCRIPCION);
                            }
                        }
                        List<PagosMasivosNuevo> ListaExc = new List<PagosMasivosNuevo>();
                        ListaExc = new List<PagosMasivosNuevo>();

                        string SocExc = string.Empty;
                        string RutExc = string.Empty;

                        Microsoft.Office.Interop.Excel._Application xlApp;
                        Microsoft.Office.Interop.Excel._Workbook xlLibro;
                        Microsoft.Office.Interop.Excel._Worksheet xlHoja1;
                        Microsoft.Office.Interop.Excel.Sheets xlHojas;
                        string fileName = txtArchivo.Text;
                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlApp.Visible = false;
                        xlLibro = xlApp.Workbooks.Open(fileName);
                        try
                        {
                            xlHojas = xlLibro.Sheets;
                            try
                            {
                                int k = 1;
                                xlHoja1 = (Microsoft.Office.Interop.Excel._Worksheet)xlHojas["PagoMasivoCliente"];
                                int n = xlHoja1.UsedRange.Rows.Count;
                                PrgBarExcel.Maximum = n;
                                int j = 4;
                                int m = 2;
                                int l = 1;
                                int verificador = 0;
                                SocExc = "";
                                RutExc = "";
                                PrgBarExcel.Value = 0;
                                string row = "";
                                string col = "";
                                string value = "";

                                PagosMasivosNuevo pagosm = new PagosMasivosNuevo(row, col, value);
                                col = "2";
                                while (col.Length < 4)
                                {
                                    col = "0" + col;
                                }
                                row = "1";
                                while (row.Length < 4)
                                {
                                    row = "0" + row;
                                }
                                value = (string)xlHoja1.Cells["1", "B"].Text;
                                pagosm.COL = col;
                                pagosm.ROW = row;
                                pagosm.VALUE = value;
                                ListaExc.Add(pagosm);
                                pagosm = new PagosMasivosNuevo(row, col, value);
                                col = "4";
                                while (col.Length < 4)
                                {
                                    col = "0" + col;
                                }
                                row = "1";
                                while (row.Length < 4)
                                {
                                    row = "0" + row;
                                }
                                value = (string)xlHoja1.Cells["1", "D"].Text;
                                pagosm.COL = col;
                                pagosm.ROW = row;
                                pagosm.VALUE = value;
                                ListaExc.Add(pagosm);
                                pagosm = new PagosMasivosNuevo(row, col, value);
                                col = "2";
                                while (col.Length < 4)
                                {
                                    col = "0" + col;
                                }
                                row = "2";
                                while (row.Length < 4)
                                {
                                    row = "0" + row;
                                }
                                value = (string)xlHoja1.Cells["2", "B"].Text;
                                pagosm.COL = col;
                                pagosm.ROW = row;
                                pagosm.VALUE = value;
                                ListaExc.Add(pagosm);
                                for (int i = 3; i <= n; i++)
                                {
                                    if (verificador >= 2)
                                    {
                                        break;
                                    }
                                    if (((string)xlHoja1.Cells[j, "A"].Text != "") && ((string)xlHoja1.Cells[j, "B"].Text != ""))
                                    {
                                        for (int r = 1; r <= 2; r++)
                                        {
                                            row = Convert.ToString(j);
                                            while (row.Length < 4)
                                            {
                                                row = "0" + row;
                                            }
                                            col = Convert.ToString(r);
                                            while (col.Length < 4)
                                            {
                                                col = "0" + col;
                                            }
                                            if (r == 1)
                                                value = (string)xlHoja1.Cells[j, "A"].Text;
                                            else
                                                value = (string)xlHoja1.Cells[j, "B"].Text;
                                            pagosm = new PagosMasivosNuevo(row, col, value);
                                            pagosm.COL = col;
                                            pagosm.ROW = row;
                                            pagosm.VALUE = value;
                                            if (value != "")
                                            {
                                                ListaExc.Add(pagosm);
                                            }
                                        }
                                        j++;
                                        PrgBarExcel.Value = (PrgBarExcel.Value + 1);
                                    }
                                    else
                                    {
                                        verificador++;

                                    }
                                }
                                SocExc = (string)xlHoja1.Cells[2, "B"].Text;
                                RutExc = (string)xlHoja1.Cells[1, "B"].Text;

                                pagosm = new PagosMasivosNuevo(row, col, value);
                                col = "2";
                                while (col.Length < 4)
                                {
                                    col = "0" + col;
                                }
                                row = Convert.ToString(xlHoja1.UsedRange.Rows.Count);
                                while (row.Length < 4)
                                {
                                    row = "0" + row;
                                }

                                value = (string)xlHoja1.Cells[Convert.ToString(xlHoja1.UsedRange.Rows.Count), "B"].Text;
                                pagosm.COL = col;
                                pagosm.ROW = row;
                                pagosm.VALUE = value;

                                decimal ValorAux2 = Convert.ToDecimal(value);
                                string PagosMasivo1 = string.Format("{0:0,0.##}", ValorAux2);
                                textBlock4Masivo.Text = Convert.ToString(PagosMasivo1);
                                txtMontoFPMasivo.Text = Convert.ToString(PagosMasivo1);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                            }
                            if (ListaExc.Count == 0)
                            {
                                System.Windows.MessageBox.Show("Error en el archivo excel a cargar. Revise formato del archivo o el formato de la plantilla o si el archivo tiene datos");
                            }
                        }
                        finally
                        {
                            xlLibro.Close(false);
                            xlApp.Quit();
                            PrgBarExcel.Value = 0;
                            GC.Collect();
                        }

                        cmbVPMedioPagMasivo.ItemsSource = VP;
                        cmbBancoPropMasivoMasivo.ItemsSource = null;
                        cmbBancoPropMasivoMasivo.Items.Clear();
                        cmbCuentasBancosPropMasivo.ItemsSource = null;
                        cmbCuentasBancosPropMasivo.Items.Clear();
                        cmbBancoMasivo.ItemsSource = null;
                        cmbBancoMasivo.Items.Clear();

                        GBViasPagoMasivos.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                        GBViasPagoMasivos.Margin = new Thickness(1, 406, 6, 0);
                        GBViasPagoMasivos.VerticalAlignment = VerticalAlignment.Top;
                        GBViasPagoMasivos.Visibility = Visibility.Visible;
                    }
                }
                else
                {

                    System.Windows.Forms.MessageBox.Show("Debe Cargar Archivo Excel ");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        //*** BOTON QUE DESPLIEGA EL GROUPBOX DE LAS FORMAS DE PAGO Y CALCULA EL MONTO DE LOS DOCUMENTOS SELECCIONADOS
        private void btnPagos_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GBViasPago.Visibility = Visibility.Collapsed;
                cmbBancoProp.ItemsSource = null;
                cmbBancoProp.Items.Clear();
                cmbCuentasBancosProp.ItemsSource = null;
                cmbCuentasBancosProp.Items.Clear();
                cmbBanco.ItemsSource = null;
                cmbBanco.Items.Clear();
                txtVuelto.Visibility = Visibility.Collapsed;
                lbvuelto.Visibility = Visibility.Collapsed;
                partidaseleccionadas.Clear();
                List<T_DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOSAUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGPagos.Items.Count > 0)
                {
                    for (int i = 0; i < DGPagos.Items.Count; i++)
                    {
                        if (i == 0)
                            DGPagos.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGPagos.Items.CurrentItem as T_DOCUMENTOSAUX);
                        }
                        DGPagos.Items.MoveCurrentToNext();
                    }
                }
                for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                {
                    if (partidaseleccionadasaux2[k].ISSELECTED == true)
                    {
                        T_DOCUMENTOS partOpen = new T_DOCUMENTOS();
                        partOpen.ACC = partidaseleccionadasaux2[k].ACC;
                        partOpen.CEBE = partidaseleccionadasaux2[k].CEBE;
                        partOpen.CLASE_CUENTA = partidaseleccionadasaux2[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = partidaseleccionadasaux2[k].CLASE_DOC;
                        partOpen.CME = partidaseleccionadasaux2[k].CME;
                        partOpen.COD_CLIENTE = partidaseleccionadasaux2[k].COD_CLIENTE;
                        partOpen.COND_PAGO = partidaseleccionadasaux2[k].COND_PAGO;
                        partOpen.CONTROL_CREDITO = partidaseleccionadasaux2[k].CONTROL_CREDITO;
                        partOpen.DIAS_ATRASO = partidaseleccionadasaux2[k].DIAS_ATRASO;
                        partOpen.ESTADO = partidaseleccionadasaux2[k].ESTADO;
                        partOpen.FECHA_DOC = partidaseleccionadasaux2[k].FECHA_DOC;
                        partOpen.FECVENCI = partidaseleccionadasaux2[k].FECVENCI;
                        partOpen.ICONO = partidaseleccionadasaux2[k].ICONO;
                        partOpen.MONEDA = partidaseleccionadasaux2[k].MONEDA;
                        partOpen.MONTO = partidaseleccionadasaux2[k].MONTO;
                        partOpen.MONTOF = partidaseleccionadasaux2[k].MONTOF;
                        partOpen.MONTO_ABONADO = partidaseleccionadasaux2[k].MONTO_ABONADO;
                        partOpen.MONTOF_ABON = partidaseleccionadasaux2[k].MONTOF_ABON;
                        partOpen.MONTO_PAGAR = partidaseleccionadasaux2[k].MONTO_PAGAR;
                        partOpen.MONTOF_PAGAR = partidaseleccionadasaux2[k].MONTOF_PAGAR;
                        partOpen.NDOCTO = partidaseleccionadasaux2[k].NDOCTO;
                        partOpen.NOMCLI = partidaseleccionadasaux2[k].NOMCLI;
                        partOpen.NREF = partidaseleccionadasaux2[k].NREF;
                        partOpen.RUTCLI = partidaseleccionadasaux2[k].RUTCLI;
                        partOpen.SOCIEDAD = partidaseleccionadasaux2[k].SOCIEDAD;
                        partidaseleccionadas.Add(partOpen);
                    }
                }
                bool permiso = true;
                if (chkAbono.IsChecked == true)
                {
                    for (int i = 0; i < partidaseleccionadas.Count; i++)
                    {
                    }
                }
                if (permiso == true)
                {
                    List<ViasPago> Condiciones = new List<ViasPago>();
                    List<string> CondicionPago = new List<string>();
                    string CondPago = "";

                    ViasPago Condic;
                    for (int i = 0; i < partidaseleccionadas.Count; i++)
                    {
                        try
                        {
                            if (partidaseleccionadas[i].COND_PAGO != "")
                            {
                                if (CondicionPago.Contains(partidaseleccionadas[i].COND_PAGO) == false)
                                {
                                    CondicionPago.Add(partidaseleccionadas[i].COND_PAGO);
                                    if (i == 0)
                                    {
                                        CondPago = partidaseleccionadas[i].COND_PAGO;
                                    }
                                    else
                                    {
                                        CondPago = CondPago + "  " + partidaseleccionadas[i].COND_PAGO;
                                    }
                                }
                            }

                            Condic = new ViasPago(partidaseleccionadas[i].ACC, partidaseleccionadas[i].COND_PAGO, partidaseleccionadas[i].CME);
                            Condic = new ViasPago(partidaseleccionadas[i].ACC, partidaseleccionadas[i].COND_PAGO, Convert.ToString(textBlock6.Content));
                            Condiciones.Add(Condic);

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message + ex.StackTrace);
                            System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                        }
                    }
                    DateTime Anual = Calendario.SelectedDate.Value;
                    String EjercicioValue = Convert.ToString(Anual.Year);
                    MatrizDePago matrizpago = new MatrizDePago();
                    String Protesto = "";
                    if (GBPagoDocs.IsVisible)
                    {
                        if (tabItem1.IsSelected)
                        {
                            Protesto = partidasabiertas.protesto;
                        }
                        if (tabItem2.IsSelected)
                        {
                            Protesto = documentospagosmasivos.protesto;
                        }
                        if (tabItem3.IsSelected)
                        {
                            Protesto = anticipos.protesto;
                        }
                    }
                    bool Excepcion = true;
                    bool Excepcion2 = true;
                    bool Excepcion3 = true;
                    string Excep = "";
                    double Monto2 = 0;
                    double Monto3 = 0;
                    double Monto4 = 0;

                    for (int k = 0; k < partidaseleccionadas.Count; k++)
                    {
                        if (partidaseleccionadas[k].MONTOF.Contains("-"))
                        {
                            partidaseleccionadas[k].MONTOF = partidaseleccionadas[k].MONTOF.Replace("-", "");
                            partidaseleccionadas[k].MONTOF = "-" + partidaseleccionadas[k].MONTOF.Trim();
                            Monto3 = Monto3 + Convert.ToDouble(partidaseleccionadas[k].MONTOF);
                        }
                        else
                        {
                            Monto4 = Monto4 + Convert.ToDouble(partidaseleccionadas[k].MONTOF_PAGAR);
                        }

                        Monto2 = Monto3 + Monto4;

                    }
                    if (Monto2 < 0)
                    {
                        Excepcion2 = false;
                    }

                    if (Monto2 == 0)
                    {
                        Excepcion3 = false;
                    }
                    //Recaudación de Servicios y Repuestos
                    if (tabItem1.IsSelected)
                    {
                        for (int k = 0; k < partidaseleccionadas.Count; k++)
                        {

                            if (partidaseleccionadas[k].NREF != "")
                            {
                                if (((partidaseleccionadas[k].COND_PAGO == "A009") | (partidaseleccionadas[k].COND_PAGO == "A010")) & (partidaseleccionadas[k].NREF.Substring(0, 1) == "E"))
                                {
                                    ;
                                }
                                else
                                {
                                    Excepcion = false;
                                }
                            }
                            else
                            {
                                if ((partidaseleccionadas[k].COND_PAGO == "A009") | (partidaseleccionadas[k].COND_PAGO == "A010"))
                                {
                                    ;
                                }
                                else
                                {
                                    Excepcion = false;
                                }
                            }
                        }
                    }
                    //Recaudacion de anticipos
                    if (tabItem3.IsSelected)
                    {
                        for (int k = 0; k < partidaseleccionadas.Count; k++)
                        {
                            if ((partidaseleccionadas[k].COND_PAGO == "A009") | (partidaseleccionadas[k].COND_PAGO == "A010"))
                            {
                                ;
                            }
                            else
                            {
                                Excepcion = false;
                            }
                        }
                    }
                    if (Excepcion == true)
                    {
                        if (Excepcion2 == false)
                        {
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "9", "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                        if (Excepcion2 == true & Monto2 != 0)
                        {
                            Excep = "X";
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Excep, "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                        if (Excepcion2 == true & Monto2 == 0)
                        {

                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "8", "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                    }
                    if (Excepcion == false)
                    {
                        if (Excepcion2 == false)
                        {
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "9", "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }

                        if (Excepcion2 == true & Monto2 != 0)
                        {
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Excep, "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }

                        if (Excepcion3 == false & Monto2 == 0)
                        {
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "8", "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                    }
                    if (matrizpago.ObjDatosViasPago.Count > 0)
                    {
                        List<string> VP = new List<string>();
                        for (int i = 0; i < matrizpago.ObjDatosViasPago.Count; i++)
                        {
                            if (VP.Contains(matrizpago.ObjDatosViasPago[i].VIA_PAGO + " - " + matrizpago.ObjDatosViasPago[i].DESCRIPCION))
                            {
                                ;
                            }
                            else
                            {
                                VP.Add(matrizpago.ObjDatosViasPago[i].VIA_PAGO + " - " + matrizpago.ObjDatosViasPago[i].DESCRIPCION);
                            }
                        }
                        cmbVPMedioPag.ItemsSource = VP;

                        int posicion = 0;
                        double Monto = 0;
                        double Monto5 = 0;

                        for (int i = 0; i < partidaseleccionadas.Count; i++)
                        {
                            try
                            {
                                partidaseleccionadas[i].MONTOF = partidaseleccionadas[i].MONTOF.Trim();
                                if (partidaseleccionadas[i].MONTOF.Contains("-"))
                                {
                                    posicion = partidaseleccionadas[i].MONTOF.IndexOf("-");
                                    if (posicion == partidaseleccionadas[i].MONTOF.Length - 1)
                                    {
                                        partidaseleccionadas[i].MONTOF = partidaseleccionadas[i].MONTOF.Substring(posicion, 1) + partidaseleccionadas[i].MONTOF.Substring(0, posicion);
                                    }
                                }
                                if (partidaseleccionadas[i].MONTOF_PAGAR == "")
                                {
                                    partidaseleccionadas[i].MONTOF = "0";
                                }

                                Monto = Monto3 + Monto4;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                            }
                        }
                        cmbBancoProp.ItemsSource = null;
                        cmbBancoProp.Items.Clear();
                        cmbCuentasBancosProp.ItemsSource = null;
                        cmbCuentasBancosProp.Items.Clear();
                        cmbBanco.ItemsSource = null;
                        cmbBanco.Items.Clear();
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            //string Valor = Convert.ToString(Monto);
                            //if (Valor.Contains("-"))
                            //{
                            //    Valor = "-" + Valor.Replace("-", "");
                            //}
                            //Valor = Valor.Replace(".", "");
                            //Valor = Valor.Replace(",", "");
                            //decimal ValorAux = Convert.ToDecimal(Valor);
                            //string monedachil = string.Format("{0:0,0}", ValorAux);

                            textBlock4.Text = Formato.FormatoMoneda(Convert.ToString(Monto));
                            txtMontoFP.Text = Formato.FormatoMoneda(Convert.ToString(Monto));
                            txtMontoFP.IsEnabled = true;
                        }
                        else
                        {
                            //string moneda = string.Format("{0:0,0.##}", Monto);
                            textBlock4.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(Monto));
                            txtMontoFP.IsEnabled = true;
                        }
                        GBViasPago.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                        GBViasPago.Margin = new Thickness(1, 406, 6, 0);
                        GBViasPago.VerticalAlignment = VerticalAlignment.Top;
                        lblCondPago.Content = CondPago;
                        GBViasPago.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No existen condiciones de pago definidos para este cliente");
                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
            }
        }

        //BOTON QUE DESPLIEGA EL GRID DE DOCUMENTOS POR PAGAR A PARTIR DE LA BUSQUEDA DE UN RUT O NUMERO DE DOCUMENTO
        private void btnBuscarP_Click(object sender, RoutedEventArgs e)
        {
            GBViasPago.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            ListaDocumentosPendientes();
            GC.Collect();
        }

        //BOTON QUE UBICA EL ARCHIVO A CARGAR EN PAGOS MASIVOS
        private void button3_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel Old files (*.xls)|*.xls|All files (*.*)|*.*";
            if (Convert.ToBoolean(openFileDialog.ShowDialog()) == true)
                txtArchivo.Text = openFileDialog.InitialDirectory;
            txtArchivo.Text = openFileDialog.FileName;
            GC.Collect();
        }
        ////BOTON QUE DESPLIEGA EL GRID DE DOCUMENTOS POR PAGAR DE MODO MASIVO A PARTIR DE LA BUSQUEDA DE UN RUT O NUMERO DE DOCUMENTO
        private void btnBuscarPM_Click(object sender, RoutedEventArgs e)
        {
            LimpiarViasDePago();
            ListaDocumentosPendientesCargasMasivas(txtArchivo.Text);
            LimpiarPagMasivo();
            GC.Collect();
        }
        //CLICK DEL BOTON QUE MUESTRA LOS DOCUMENTOS PARA REALIZAR ANTICIPOS  
        private void button5_Click(object sender, RoutedEventArgs e)
        {
            GBViasPago.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            ListaDocumentosPendientesAnticipos();
            DPFechActual.Text = Calendario.Text;
            cheques.Clear();
            GC.Collect();
        }
        //CLICK DEL BOTON DE LA BARRA DE HERRAMIENTAS INICIO
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            GBInicio.Visibility = Visibility.Visible;
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            btnAutAnul.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            GC.Collect();
        }
        //BOTON PARA LA EMISION DE LAS NOTAS DE CREDITO (NC)
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            CargarDatos();
            PagoDocumentos.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBViasPagoMasivos.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            GC.Collect();
        }
        //CLICK  QUE INGRESA LOS MONTOS DE LOS MEDIOS DE PAGOS EN LA GRILLA TOTALIZADORA Pagos Masivos
        private void btnAgregaMtoMasivo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //InitializeComponent();
                if (cmbVPMedioPagMasivo.Text != "")
                {
                    string MedioPago = cmbVPMedioPagMasivo.Text as string;
                    MedioPago = MedioPago.Substring(0, 1);

                    switch (MedioPago)
                    {
                        case "1": //Documentos tributarios
                            {
                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                break;
                            }
                        case "9": //Saldo a favor del cliente
                            {
                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                break;
                            }
                        case "8": //Compensacion anticipo saldo 0 
                            {
                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                break;
                            }
                        case "K": //Carta curse
                            {

                                if (DPFechActualMasivo.Text != "")
                                {
                                    if (DPFechVencMasivo.Text != "")
                                    {
                                        if (txtNumDocMasivo.Text != "")
                                        {
                                            if (txtMontoFPMasivo.Text != "")
                                            {
                                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el monto del pago por carta curse");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el número de la carta curse");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión");
                                }
                                break;
                            }
                        case "F": //Cheque a fecha
                            {
                                if (txtCodAutMasivo.Text != "")
                                {
                                    if (DPFechVencMasivo.Text != "")
                                    {
                                        if (cmbBancoMasivo.Text != "")
                                        {
                                            if (txtSucursalMasivo.Text != "")
                                            {
                                                if (txtNumDocMasivo.Text != "")
                                                {
                                                    if (txtMontoFPMasivo.Text != "")
                                                    {
                                                        if (txtNumCuentaMasivo.Text != "")
                                                        {
                                                            if (txtCantDocMasivo.Text != "")
                                                            {
                                                                if (cmbIntervaloMasivo.Text != "")
                                                                {
                                                                    if (txtRUTGiradorMasivo.Text != "")
                                                                    {
                                                                        String RUT = DigitoVerificador(txtRUTGiradorMasivo.Text.ToUpper());
                                                                        if (txtRUTGiradorMasivo.Text.ToUpper() != "")
                                                                        {
                                                                            if (RUT != txtRUTGiradorMasivo.Text.ToUpper())
                                                                            {
                                                                                System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                                                                                txtRUTGiradorMasivo.Focus();
                                                                            }
                                                                            else
                                                                            {
                                                                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        System.Windows.MessageBox.Show("Ingrese el Rut del girador");
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    System.Windows.MessageBox.Show("Ingrese el intervalo de días");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                System.Windows.MessageBox.Show("Ingrese el número de cuotas");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("Ingrese el número de la cuenta bancaria");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        System.Windows.MessageBox.Show("Ingrese el monto del cheque");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Windows.MessageBox.Show("Ingrese el número de cheque");
                                                }
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese la plaza");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el banco emisor");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento del cheque");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el Codigo Autorización");
                                }
                                break;
                            }
                        case "G": //Cheque al día
                            {
                                if (txtCodAutMasivo.Text != "")
                                {
                                    if (cmbBancoMasivo.Text != "")
                                    {
                                        if (txtSucursalMasivo.Text != "")
                                        {
                                            if (txtNumDocMasivo.Text != "")
                                            {
                                                if (txtMontoFPMasivo.Text != "")
                                                {
                                                    if (txtNumCuentaMasivo.Text != "")
                                                    {
                                                        if (txtRUTGiradorMasivo.Text != "")
                                                        {
                                                            String RUT = DigitoVerificador(txtRUTGiradorMasivo.Text.ToUpper());
                                                            if (txtRUTGiradorMasivo.Text.ToUpper() != "")
                                                            {
                                                                if (RUT != txtRUTGiradorMasivo.Text.ToUpper())
                                                                {
                                                                    System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                                                                    txtRUTGiradorMasivo.Focus();
                                                                }
                                                                else
                                                                {
                                                                    IngresoFormasDePagoYMontosMasivos(MedioPago);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("Ingrese el Rut del girador");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        System.Windows.MessageBox.Show("Ingrese el número de la cuenta bancaria");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Windows.MessageBox.Show("Ingrese el monto del cheque");
                                                }
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el número de cheque");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese la plaza");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el banco emisor");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el Codigo Autorizacion");
                                }

                                break;
                            }
                        case "M": //Contrato compra-venta
                            {
                                if (txtMontoFPMasivo.Text != "")
                                {
                                    IngresoFormasDePagoYMontosMasivos(MedioPago);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el monto del pago por contrato");
                                }
                                break;
                            }
                        case "D": //Deposito a plazo
                            {
                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                break;
                            }
                        case "B": //Deposito en cliente corriente
                            {
                                if (DPFechActualMasivo.Text != "")
                                {
                                    //if (cmbBanco.Text != "")
                                    //{
                                    if (txtNumDocMasivo.Text != "")
                                    {
                                        if (txtMontoFPMasivo.Text != "")
                                        {
                                            IngresoFormasDePagoYMontosMasivos(MedioPago);
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el monto del pago por depósito en cuenta corriente");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el número de depósito en cuenta corriente");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión del depósito en cuenta corriente");
                                }
                                //}
                                //else
                                //{
                                //    System.Windows.MessageBox.Show("Ingrese el banco del depósito en cuenta corriente");
                                //}
                                break;
                            }
                        case "L": //Letras
                            {
                                if (txtNumDocMasivo.Text != "")
                                {
                                    IngresoFormasDePagoYMontosMasivos(MedioPago);

                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la letra");
                                }
                                break;
                            }
                        case "P": //Pagaré
                            {
                                if (DPFechVencMasivo.Text != "")
                                {
                                    if (txtNumDocMasivo.Text != "")
                                    {
                                        if (txtMontoFPMasivo.Text != "")
                                        {
                                            IngresoFormasDePagoYMontosMasivos(MedioPago);
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el monto del pago por pagaré");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el número del pagaré");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento del pagaré");
                                }
                                break;
                            }
                        case "E": //Pago en efectivo
                            {
                                if (txtMontoFPMasivo.Text != "")
                                {
                                    IngresoFormasDePagoYMontosMasivos(MedioPago);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el monto del pago en efectivo");
                                }
                                break;
                            }
                        case "S": //Tarjeta de crédito
                            {
                                if (txtNumDocMasivo.Text != "")
                                {
                                    if (txtMontoFPMasivo.Text != "")
                                    {
                                        if (txtCodAutMasivo.Text != "")
                                        {
                                            if (txtCodOpMasivo.Text != "")
                                            {
                                                //if (txtAsig.Text != "")
                                                //{
                                                if (cmbTipoTarjetaMasivo.Text != "")
                                                {
                                                    if (txtCantDocMasivo.Text != "")
                                                    {
                                                        if (cmbIntervaloMasivo.Text != "")
                                                        {
                                                            IngresoFormasDePagoYMontosMasivos(MedioPago);
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("Ingrese el intervalo de días");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        System.Windows.MessageBox.Show("Ingrese el número de cuotas");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Windows.MessageBox.Show("Ingrese el tipo de tarjeta");
                                                }
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el código de operación por tarjeta");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el código de autorización por tarjeta");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago por tarjeta");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la tarjeta de crédito");
                                }
                                break;
                            }
                        case "R": //Tarjeta de débito
                            {
                                if (txtNumDocMasivo.Text != "")
                                {
                                    if (txtMontoFPMasivo.Text != "")
                                    {
                                        if (txtCodAutMasivo.Text != "")
                                        {
                                            if (txtCodOpMasivo.Text != "")
                                            {
                                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el código de operación por tarjeta");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el código de autorización por tarjeta");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago por tarjeta");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la tarjeta de débito");
                                }
                                break;
                            }
                        case "N": //Servipag
                            {
                                if (txtNumDocMasivo.Text != "")
                                {
                                    if (txtMontoFPMasivo.Text != "")
                                    {
                                        if (txtCodAutMasivo.Text != "")
                                        {
                                            if (txtCodOpMasivo.Text != "")
                                            {
                                                IngresoFormasDePagoYMontosMasivos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el código de operación");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el código de autorización");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la Transacción");
                                }
                                break;
                            }
                        case "U": //Transferencia bancaria
                            {
                                if (DPFechActualMasivo.Text != "")
                                {
                                    if (cmbBancoPropMasivoMasivo.Text != "")
                                    {
                                        IngresoFormasDePagoYMontosMasivos(MedioPago);
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el banco propio para la transferencia bancaria");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión de la transferencia bancaria");
                                }
                                break;
                            }
                        case "V": //Vale vista recibido
                            {

                                if (DPFechActualMasivo.Text != "")
                                {

                                    if (DPFechVencMasivo.Text != "")
                                    {
                                        IngresoFormasDePagoYMontosMasivos(MedioPago);
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento del vale vista");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión del vale vista");
                                }
                                break;
                            }
                        case "A": //Vehiculo en parte de pago
                            {
                                if (txtNumDocMasivo.Text != "")
                                {
                                    if (txtPatenteMasivo.Text != "")
                                    {
                                        if (txtMontoFPMasivo.Text != "")
                                        {
                                            IngresoFormasDePagoYMontosMasivos(MedioPago);
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el monto del pago");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el número de la patente del vehículo");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de via de pago");
                                }
                                break;

                            }
                        default:
                            {
                                break;
                            }
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Ingrese el tipo de documento (forma de pago)");
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        //CLICK  QUE INGRESA LOS MONTOS DE LOS MEDIOS DE PAGOS EN LA GRILLA TOTALIZADORA 
        private void btnAgregaMto_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //InitializeComponent();
                if (cmbVPMedioPag.Text != "")
                {
                    string MedioPago = cmbVPMedioPag.Text as string;
                    MedioPago = MedioPago.Substring(0, 1);

                    switch (MedioPago)
                    {

                        case "1": //Documentos tributarios
                            {
                                IngresoFormasDePagoYMontos(MedioPago);
                                break;
                            }
                        case "9": //Saldo a favor del cliente
                            {
                                IngresoFormasDePagoYMontos(MedioPago);
                                break;
                            }
                        case "8": //Compensacion anticipo saldo 0 
                            {
                                IngresoFormasDePagoYMontos(MedioPago);
                                break;
                            }
                        case "K": //Carta curse
                            {

                                if (DPFechActual.Text != "")
                                {
                                    if (DPFechVenc.Text != "")
                                    {
                                        if (txtNumDoc.Text != "")
                                        {
                                            if (txtMontoFP.Text != "")
                                            {
                                                IngresoFormasDePagoYMontos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el monto del pago por carta curse");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el número de la carta curse");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión");
                                }
                                break;
                            }
                        case "F": //Cheque a fecha
                            {
                                if (txtCodAut.Text != "")
                                {
                                    if (DPFechVenc.Text != "")
                                    {
                                        if (cmbBanco.Text != "")
                                        {
                                            if (txtSucursal.Text != "")
                                            {
                                                if (txtNumDoc.Text != "")
                                                {
                                                    if (txtMontoFP.Text != "")
                                                    {
                                                        if (txtNumCuenta.Text != "")
                                                        {
                                                            if (txtCantDoc.Text != "")
                                                            {
                                                                if (cmbIntervalo.Text != "")
                                                                {
                                                                    if (txtRUTGirador.Text != "")
                                                                    {
                                                                        String RUT = DigitoVerificador(txtRUTGirador.Text.ToUpper());
                                                                        if (txtRUTGirador.Text.ToUpper() != "")
                                                                        {
                                                                            if (RUT != txtRUTGirador.Text.ToUpper())
                                                                            {
                                                                                System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                                                                                txtRUTGirador.Focus();
                                                                            }
                                                                            else
                                                                            {
                                                                                IngresoFormasDePagoYMontos(MedioPago);
                                                                            }
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        System.Windows.MessageBox.Show("Ingrese el Rut del girador");
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    System.Windows.MessageBox.Show("Ingrese el intervalo de días");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                System.Windows.MessageBox.Show("Ingrese el número de cuotas");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("Ingrese el número de la cuenta bancaria");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        System.Windows.MessageBox.Show("Ingrese el monto del cheque");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Windows.MessageBox.Show("Ingrese el número de cheque");
                                                }
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese la plaza");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el banco emisor");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento del cheque");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el Codigo Autorización");
                                }
                                break;
                            }
                        case "G": //Cheque al día
                            {
                                if (txtCodAut.Text != "")
                                {
                                    if (cmbBanco.Text != "")
                                    {
                                        if (txtSucursal.Text != "")
                                        {
                                            if (txtNumDoc.Text != "")
                                            {
                                                if (txtMontoFP.Text != "")
                                                {
                                                    if (txtNumCuenta.Text != "")
                                                    {
                                                        if (txtRUTGirador.Text != "")
                                                        {
                                                            String RUT = DigitoVerificador(txtRUTGirador.Text.ToUpper());
                                                            if (txtRUTGirador.Text.ToUpper() != "")
                                                            {
                                                                if (RUT != txtRUTGirador.Text.ToUpper())
                                                                {
                                                                    System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                                                                    txtRUTGirador.Focus();
                                                                }
                                                                else
                                                                {
                                                                    IngresoFormasDePagoYMontos(MedioPago);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("Ingrese el Rut del girador");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        System.Windows.MessageBox.Show("Ingrese el número de la cuenta bancaria");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Windows.MessageBox.Show("Ingrese el monto del cheque");
                                                }
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el número de cheque");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese la plaza");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el banco emisor");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el Codigo Autorizacion");
                                }

                                break;
                            }
                        case "M": //Contrato compra-venta
                            {
                                if (txtMontoFP.Text != "")
                                {
                                    IngresoFormasDePagoYMontos(MedioPago);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el monto del pago por contrato");
                                }
                                break;
                            }
                        case "D": //Deposito a plazo
                            {
                                IngresoFormasDePagoYMontos(MedioPago);
                                break;
                            }
                        case "B": //Deposito en cliente corriente
                            {
                                if (DPFechActual.Text != "")
                                {
                                    //if (cmbBanco.Text != "")
                                    //{
                                    if (txtNumDoc.Text != "")
                                    {
                                        if (txtMontoFP.Text != "")
                                        {
                                            IngresoFormasDePagoYMontos(MedioPago);
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el monto del pago por depósito en cuenta corriente");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el número de depósito en cuenta corriente");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión del depósito en cuenta corriente");
                                }
                                break;
                            }
                        case "L": //Letras
                            {
                                if (txtNumDoc.Text != "")
                                {
                                    IngresoFormasDePagoYMontos(MedioPago);

                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la letra");
                                }
                                break;
                            }
                        case "P": //Pagaré
                            {
                                if (DPFechVenc.Text != "")
                                {
                                    if (txtNumDoc.Text != "")
                                    {
                                        if (txtMontoFP.Text != "")
                                        {
                                            IngresoFormasDePagoYMontos(MedioPago);
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el monto del pago por pagaré");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el número del pagaré");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento del pagaré");
                                }
                                break;
                            }
                        case "E": //Pago en efectivo
                            {
                                if (txtMontoFP.Text != "")
                                {
                                    IngresoFormasDePagoYMontos(MedioPago);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el monto del pago en efectivo");
                                }
                                break;
                            }
                        case "H": //Pago Dolares    
                            {
                                if (txtMontoFP.Text != "")
                                {
                                    IngresoFormasDePagoYMontos(MedioPago);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el monto del pago en Dolares");
                                }
                                break;
                            }
                        case "J": //Pago en Euros
                            {
                                if (txtMontoFP.Text != "")
                                {
                                    IngresoFormasDePagoYMontos(MedioPago);
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el monto del pago en Euros");
                                }
                                break;
                            }
                        case "S": //Tarjeta de crédito
                            {
                                if (txtNumDoc.Text != "")
                                {
                                    if (txtMontoFP.Text != "")
                                    {
                                        if (txtCodAut.Text != "")
                                        {
                                            if (txtCodOp.Text != "")
                                            {
                                                if (cmbTipoTarjeta.Text != "")
                                                {
                                                    if (txtCantDoc.Text != "")
                                                    {
                                                        if (cmbIntervalo.Text != "")
                                                        {
                                                            IngresoFormasDePagoYMontos(MedioPago);
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("Ingrese el intervalo de días");
                                                        }
                                                    }
                                                    else
                                                    {
                                                        System.Windows.MessageBox.Show("Ingrese el número de cuotas");
                                                    }
                                                }
                                                else
                                                {
                                                    System.Windows.MessageBox.Show("Ingrese el tipo de tarjeta");
                                                }
                    
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el código de operación por tarjeta");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el código de autorización por tarjeta");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago por tarjeta");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la tarjeta de crédito");
                                }
                                break;
                            }
                        case "R": //Tarjeta de débito
                            {
                                if (txtNumDoc.Text != "")
                                {
                                    if (txtMontoFP.Text != "")
                                    {
                                        if (txtCodAut.Text != "")
                                        {
                                            if (txtCodOp.Text != "")
                                            {
                                                IngresoFormasDePagoYMontos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el código de operación por tarjeta");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el código de autorización por tarjeta");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago por tarjeta");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la tarjeta de débito");
                                }
                                break;
                            }
                        case "I": //GIFTCARD
                            {
                                if (txtNumDoc.Text != "")
                                {
                                    if (txtMontoFP.Text != "")
                                    {
                                        if (txtCodAut.Text != "")
                                        {
                                            if (txtCodOp.Text != "")
                                            {
                                                IngresoFormasDePagoYMontos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese el código de operación por tarjeta");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el código de autorización por tarjeta");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago por tarjeta");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la tarjeta de descuento");
                                }
                                break;
                            }
                        case "N": //Servipag
                            {
                                if (txtNumDoc.Text != "")
                                {
                                    if (txtMontoFP.Text != "")
                                    {
                                        //if (txtCodAut.Text != "")
                                        //{
                                            //if (txtCodOp.Text != "")
                                            //{
                                                IngresoFormasDePagoYMontos(MedioPago);
                                            //}
                                            //else
                                            //{
                                            //    System.Windows.MessageBox.Show("Ingrese el código de operación");
                                            //}
                                        //}
                                        //else
                                        //{
                                        //    System.Windows.MessageBox.Show("Ingrese el código de autorización");
                                        //}
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el monto del pago");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de la Transacción");
                                }
                                break;
                            }
                        case "U": //Transferencia bancaria
                            {
                                if (DPFechActual.Text != "")
                                {
                                    if (cmbBancoProp.Text != "")
                                    {
                                        IngresoFormasDePagoYMontos(MedioPago);
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el banco propio para la transferencia bancaria");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión de la transferencia bancaria");
                                }
                                break;
                            }
                        case "O": //Transferencia bancaria Dolares
                            {
                                if (DPFechActual.Text != "")
                                {
                                    if (cmbBancoProp.Text != "")
                                    {
                                        IngresoFormasDePagoYMontos(MedioPago);
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el banco propio para la transferencia bancaria en dolares");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la fecha de emisión de la transferencia bancaria");
                                }
                                break;
                            }
                        case "V": //Vale vista recibido
                            {
                                if (txtSucursal.Text != "")
                                {
                                    if (cmbBanco.Text != "")
                                    {
                                        if (DPFechActual.Text != "")
                                        {

                                            if (DPFechVenc.Text != "")
                                            {
                                                IngresoFormasDePagoYMontos(MedioPago);
                                            }
                                            else
                                            {
                                                System.Windows.MessageBox.Show("Ingrese la fecha de vencimiento del vale vista");
                                            }
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese la fecha de emisión del vale vista");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Debe seleccionar banco de destino");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese la plaza");
                                }
                                break;
                            }
                        case "A": //Vehiculo en parte de pago
                            {
                                if (txtNumDoc.Text != "")
                                {
                                    if (txtPatente.Text != "")
                                    {
                                        if (txtMontoFP.Text != "")
                                        {
                                            IngresoFormasDePagoYMontos(MedioPago);
                                        }
                                        else
                                        {
                                            System.Windows.MessageBox.Show("Ingrese el monto del pago");
                                        }
                                    }
                                    else
                                    {
                                        System.Windows.MessageBox.Show("Ingrese el número de la patente del vehículo");
                                    }
                                }
                                else
                                {
                                    System.Windows.MessageBox.Show("Ingrese el número de via de pago");
                                }
                                break;

                            }
                        default:
                            {
                                break;
                            }
                      }
                }
                else
                {
                    System.Windows.MessageBox.Show("Ingrese el tipo de documento (forma de pago)");
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }

        }
        // Pagos Masivos 
        private void btnResPagosMasivo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<DetalleViasPago> ListViasPagos = new List<DetalleViasPago>();
                for (int i = 1; i <= DGCheque.Items.Count; i++)
                {
                    if (i == 1)
                    {
                        DGCheque.Items.MoveCurrentToFirst();
                    }
                    ListViasPagos.Add(DGCheque.Items.CurrentItem as DetalleViasPago);
                    DGCheque.Items.MoveCurrentToNext();
                }

                //LLAMADA AL FORM DE RESUMEN DE PAGOS
                ResumenViasPago frm = new ResumenViasPago();
                frm.DGResumenViasPago.ItemsSource = ListViasPagos;
                frm.Owner = this;
                frm.Show();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }

        }

        //CLICK QUE MUESTRA EL FORM DE RESUMEN DE PAGOS CON TODA LA INFORMACION DE LAS VIAS DE PAGO 
        private void btnResPagos_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<DetalleViasPago> ListViasPagos = new List<DetalleViasPago>();
                for (int i = 1; i <= DGCheque.Items.Count - 1; i++)
                {
                    if (i == 1)
                    {
                        DGCheque.Items.MoveCurrentToFirst();
                    }
                    ListViasPagos.Add(DGCheque.Items.CurrentItem as DetalleViasPago);
                    DGCheque.Items.MoveCurrentToNext();
                }

                //LLAMADA AL FORM DE RESUMEN DE PAGOS
                ResumenViasPago frm = new ResumenViasPago();

                frm.DGResumenViasPago.ItemsSource = ListViasPagos;
                //frm.DialogResult = true;// = this;
                frm.Owner = this;
                frm.Show();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }

        private string RemoveSpecialCharacters(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_.- ]+", "", RegexOptions.Compiled);
        }

        private void btnConfirPagMasivo_Click(object sender, RoutedEventArgs e)
        {
            //*RFC PAGO DE DOCUMENTOS
            List<DetalleViasPago> ListViasPagos = new List<DetalleViasPago>();

            for (int i = 1; i <= DGCheque.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGCheque.Items.MoveCurrentToFirst();
                }
                ListViasPagos.Add(DGCheque.Items.CurrentItem as DetalleViasPago);

                DGCheque.Items.MoveCurrentToNext();
            }

            List<T_DOCUMENTOS> DocsAPagar = new List<T_DOCUMENTOS>();
            List<T_DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOSAUX>();
            partidaseleccionadasaux2.Clear();
            if (this.DGPagos.Items.Count > 0)
            {
                for (int i = 0; i < DGPagos.Items.Count; i++)
                {
                    if (i == 0)
                        DGPagos.Items.MoveCurrentToFirst();
                    {
                        partidaseleccionadasaux2.Add(DGPagos.Items.CurrentItem as T_DOCUMENTOSAUX);
                    }
                    DGPagos.Items.MoveCurrentToNext();
                }
            }
            for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
            {
                if (partidaseleccionadasaux2[k].ISSELECTED == true)
                {
                    T_DOCUMENTOS partOpen = new T_DOCUMENTOS();
                    partOpen.ACC = partidaseleccionadasaux2[k].ACC;
                    partOpen.CEBE = partidaseleccionadasaux2[k].CEBE;
                    partOpen.CLASE_CUENTA = partidaseleccionadasaux2[k].CLASE_CUENTA;
                    partOpen.CLASE_DOC = partidaseleccionadasaux2[k].CLASE_DOC;
                    partOpen.CME = partidaseleccionadasaux2[k].CME;
                    partOpen.COD_CLIENTE = partidaseleccionadasaux2[k].COD_CLIENTE;
                    partOpen.COND_PAGO = partidaseleccionadasaux2[k].COND_PAGO;
                    partOpen.CONTROL_CREDITO = partidaseleccionadasaux2[k].CONTROL_CREDITO;
                    partOpen.DIAS_ATRASO = partidaseleccionadasaux2[k].DIAS_ATRASO;
                    partOpen.ESTADO = partidaseleccionadasaux2[k].ESTADO;
                    partOpen.FECHA_DOC = partidaseleccionadasaux2[k].FECHA_DOC;
                    partOpen.FECVENCI = partidaseleccionadasaux2[k].FECVENCI;
                    partOpen.ICONO = partidaseleccionadasaux2[k].ICONO;
                    partOpen.MONEDA = partidaseleccionadasaux2[k].MONEDA;
                    partOpen.MONTO = partidaseleccionadasaux2[k].MONTO;
                    partOpen.MONTOF = partidaseleccionadasaux2[k].MONTOF;
                    partOpen.MONTO_ABONADO = partidaseleccionadasaux2[k].MONTO_ABONADO;
                    partOpen.MONTOF_ABON = partidaseleccionadasaux2[k].MONTOF_ABON;
                    partOpen.MONTO_PAGAR = partidaseleccionadasaux2[k].MONTOF_PAGAR;
                    partOpen.MONTOF_PAGAR = partidaseleccionadasaux2[k].MONTOF_PAGAR;
                    partOpen.NDOCTO = partidaseleccionadasaux2[k].NDOCTO;
                    partOpen.NOMCLI = partidaseleccionadasaux2[k].NOMCLI;
                    partOpen.NREF = partidaseleccionadasaux2[k].NREF;
                    partOpen.RUTCLI = partidaseleccionadasaux2[k].RUTCLI;
                    partOpen.SOCIEDAD = partidaseleccionadasaux2[k].SOCIEDAD;
                    DocsAPagar.Add(partOpen);
                }
            }
            Int64 Ingreso = 0;
            for (int i = 0; i < ListViasPagos.Count; i++)
            {
                {
                    Ingreso = Ingreso + Convert.ToInt64(ListViasPagos[i].MONTO);
                }
            }
            double APagar2 = 0;
            double APagar = 0;
            for (int i = 0; i < DocsAPagar.Count; i++)
            {
                {
                    string s = "";
                    int Pos = DocsAPagar[i].MONTO.IndexOf(",");
                    if (Pos != -1)
                    {
                        s = DocsAPagar[i].MONTO.Remove(Pos, 1);
                    }
                    else
                    {
                        s = DocsAPagar[i].MONTO.Trim();
                    }
                    APagar = APagar + Convert.ToDouble(s);
                }
            }
            bool pago = false;
            try
            {
                if (tabItem1.IsSelected)
                {
                    pagodocumentosingreso.pagodocumentosingreso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content)
                        , txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text
                        , txtIdioma.Text, Convert.ToString(lblSociedad.Content), ListViasPagos, DocsAPagar
                        , Convert.ToString(lblPais.Content), cmbMoneda.Text, Convert.ToString(textBlock6.Content)
                        , Convert.ToString(textBlock7.Content), Convert.ToString(textBlock3.Text), Convert.ToString(APagar));

                    string Mensaje = "";
                    for (int i = 0; i < pagodocumentosingreso.T_Retorno.Count; i++)
                    {
                        Mensaje = Mensaje + " - " + pagodocumentosingreso.T_Retorno[i].MESSAGE + " - " + pagodocumentosingreso.T_Retorno[i].MESSAGE_V1;
                    }
                    if (pagodocumentosingreso.message != "")
                    {
                        System.Windows.MessageBox.Show(pagodocumentosingreso.message);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show(pagodocumentosingreso.pagomessage);
                    }
                    logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);

                    if (pagodocumentosingreso.comprobante != "")
                    {
                        ImpresionesDeDocumentosAutomaticas(pagodocumentosingreso.comprobante, "X");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No se generó comprobante de pago");
                    }
                    DocsAPagar.Clear();
                    ListViasPagos.Clear();

                    pagodocumentosingreso.T_Retorno.Clear();
                }
                if (tabItem3.IsSelected)
                {
                    pagoanticipos.pagoanticiposingreso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblSociedad.Content), ListViasPagos, DocsAPagar, Convert.ToString(lblPais.Content), cmbMoneda.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content), Convert.ToString(APagar2), Convert.ToString(APagar2), "");
                    string Mensaje = "";
                    if (pagoanticipos.message != "")
                    {
                        System.Windows.MessageBox.Show(pagoanticipos.message);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show(pagoanticipos.status);
                    }

                    logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);
                    if (pagoanticipos.comprobante != "")
                    {
                        ImpresionesDeDocumentosAutomaticas(pagoanticipos.comprobante, "X");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No se generó comprobante de pago");
                    }

                    pagoanticipos.T_Retorno.Clear();
                }
                LimpiarViasDePago();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }

        //BOTON QUE REALIZA EL PAGO Y CREACION DEL COMPROBANTE DE INGRESO DE NOTAS DE VENTAS Y PAGO DE ANTICIPOS
        private void btnConfirPag_Click(object sender, RoutedEventArgs e)
        {
            List<DetalleViasPago> ListViasPagos = new List<DetalleViasPago>();
            DetalleViasPago ObjPagos = new DetalleViasPago();

            for (int i = 1; i <= DGCheque.Items.Count - 1; i++)
            {
                if (i == 1)
                {
                    DGCheque.Items.MoveCurrentToFirst();
                }

                ListViasPagos.Add(DGCheque.Items.CurrentItem as DetalleViasPago);

                DGCheque.Items.MoveCurrentToNext();
            }
            if (txtVuelto.Text != "" && ListViasPagos[0].MONEDA != "CLP")
            {
                if (txtVuelto.Text != "0")
                {
                    DetalleViasPago objVuelto = new DetalleViasPago();
                    DetalleViasPago ObjOtraMoneda = new DetalleViasPago();

                    objVuelto.FECHA_EMISION = ListViasPagos[0].FECHA_EMISION;
                    objVuelto.FECHA_VENC = ListViasPagos[0].FECHA_VENC;
                    objVuelto.LAND = Convert.ToString(lblPais.Content);
                    objVuelto.MONEDA = cmbMoneda.Text;
                    objVuelto.MONTO = (Convert.ToDouble(txtVuelto.Text));
                    objVuelto.VIA_PAGO = "3";
                    ListViasPagos.Add(objVuelto);
                }
            }
            List<T_DOCUMENTOS> DocsAPagar = new List<T_DOCUMENTOS>();
            List<T_DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOSAUX>();
            partidaseleccionadasaux2.Clear();
            if (this.DGPagos.Items.Count > 0)
            {
                for (int i = 0; i < DGPagos.Items.Count; i++)
                {
                    if (i == 0)
                        DGPagos.Items.MoveCurrentToFirst();
                    {
                        partidaseleccionadasaux2.Add(DGPagos.Items.CurrentItem as T_DOCUMENTOSAUX);
                    }
                    DGPagos.Items.MoveCurrentToNext();
                }
            }
            for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
            {
                if (partidaseleccionadasaux2[k].ISSELECTED == true)
                {
                    T_DOCUMENTOS partOpen = new T_DOCUMENTOS();
                    partOpen.ACC = partidaseleccionadasaux2[k].ACC;
                    partOpen.CEBE = partidaseleccionadasaux2[k].CEBE;
                    partOpen.CLASE_CUENTA = partidaseleccionadasaux2[k].CLASE_CUENTA;
                    partOpen.CLASE_DOC = partidaseleccionadasaux2[k].CLASE_DOC;
                    partOpen.CME = partidaseleccionadasaux2[k].CME;
                    partOpen.COD_CLIENTE = partidaseleccionadasaux2[k].COD_CLIENTE;
                    partOpen.COND_PAGO = partidaseleccionadasaux2[k].COND_PAGO;
                    partOpen.CONTROL_CREDITO = partidaseleccionadasaux2[k].CONTROL_CREDITO;
                    partOpen.DIAS_ATRASO = partidaseleccionadasaux2[k].DIAS_ATRASO;
                    partOpen.ESTADO = partidaseleccionadasaux2[k].ESTADO;
                    partOpen.FECHA_DOC = partidaseleccionadasaux2[k].FECHA_DOC;
                    partOpen.FECVENCI = partidaseleccionadasaux2[k].FECVENCI;
                    partOpen.ICONO = partidaseleccionadasaux2[k].ICONO;
                    partOpen.MONEDA = partidaseleccionadasaux2[k].MONEDA;
                    partOpen.MONTO = partidaseleccionadasaux2[k].MONTO;
                    partOpen.MONTOF = partidaseleccionadasaux2[k].MONTOF;
                    partOpen.MONTO_ABONADO = partidaseleccionadasaux2[k].MONTO_ABONADO;
                    partOpen.MONTOF_ABON = partidaseleccionadasaux2[k].MONTOF_ABON;
                    partOpen.MONTO_PAGAR = partidaseleccionadasaux2[k].MONTOF_PAGAR;
                    partOpen.MONTOF_PAGAR = partidaseleccionadasaux2[k].MONTOF_PAGAR;
                    partOpen.NDOCTO = partidaseleccionadasaux2[k].NDOCTO;
                    partOpen.NOMCLI = partidaseleccionadasaux2[k].NOMCLI;
                    partOpen.NREF = partidaseleccionadasaux2[k].NREF;
                    partOpen.RUTCLI = partidaseleccionadasaux2[k].RUTCLI;
                    partOpen.SOCIEDAD = partidaseleccionadasaux2[k].SOCIEDAD;
                    DocsAPagar.Add(partOpen);
                }
            }
            Int64 Ingreso = 0;
            for (int i = 0; i < ListViasPagos.Count; i++)
            {
                {
                    Ingreso = Ingreso + Convert.ToInt64(ListViasPagos[i].MONTO);
                }
            }
            double APagar2 = 0;
            double APagar = 0;
            for (int i = 0; i < DocsAPagar.Count; i++)
            {
                {
                    string s = "";
                    int Pos = DocsAPagar[i].MONTO.IndexOf(",");
                    if (Pos != -1)
                    {
                        s = DocsAPagar[i].MONTO.Remove(Pos, 1);
                    }
                    else
                    {
                        s = DocsAPagar[i].MONTO.Trim();
                    }
                    APagar = APagar + Convert.ToDouble(s);
                }
            }
            bool pago = false;
            try
            {
                if (tabItem1.IsSelected)
                {
                    //*RFC PAGO DE DOCUMENTOS
                    pagodocumentosingreso.pagodocumentosingreso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content)
                        , txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text
                        , txtIdioma.Text, Convert.ToString(lblSociedad.Content), ListViasPagos, DocsAPagar
                        , Convert.ToString(lblPais.Content), cmbMoneda.Text, Convert.ToString(textBlock6.Content)
                        , Convert.ToString(textBlock7.Content), Convert.ToString(textBlock3.Text), Convert.ToString(APagar));

                    string Mensaje = "";
                    for (int i = 0; i < pagodocumentosingreso.T_Retorno.Count; i++)
                    {
                        Mensaje = Mensaje + " - " + pagodocumentosingreso.T_Retorno[i].MESSAGE + " - " + pagodocumentosingreso.T_Retorno[i].MESSAGE_V1;
                    }
                    if (pagodocumentosingreso.message != "")
                    {
                        System.Windows.MessageBox.Show(pagodocumentosingreso.message);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show(pagodocumentosingreso.pagomessage);
                    }
                    logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);

                    if (pagodocumentosingreso.comprobante != "")
                    {
                        ImpresionesDeDocumentosAutomaticas(pagodocumentosingreso.comprobante, "X");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No se generó comprobante de pago");
                    }
                    DocsAPagar.Clear();
                    ListViasPagos.Clear();

                    pagodocumentosingreso.T_Retorno.Clear();
                }
                if (tabItem3.IsSelected)
                {
                    //*RFC PAGO DE ANTICIPOS
                    pagoanticipos.pagoanticiposingreso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblSociedad.Content), ListViasPagos, DocsAPagar, Convert.ToString(lblPais.Content), cmbMoneda.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content), Convert.ToString(APagar2), Convert.ToString(APagar2), "");
                    string Mensaje = "";

                    if (pagoanticipos.message != "")
                    {
                        System.Windows.MessageBox.Show(pagoanticipos.message);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show(pagoanticipos.status);
                    }

                    logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);
                    if (pagoanticipos.comprobante != "")
                    {
                        ImpresionesDeDocumentosAutomaticas(pagoanticipos.comprobante, "X");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No se generó comprobante de pago");
                    }

                    pagoanticipos.T_Retorno.Clear();
                }
                LimpiarViasDePago();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        //BOTON QUE LIMPIA LOS DATOS DE LA GRILLA DE RESUMEN DE VIAS DE PAGO Y CANTIDAD y TOTALES POR DOCUMENTOS Y A PAGAR
        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            DGCheque.ItemsSource = null;
            DGCheque.Items.Clear();
            DGMediosDePagos.ItemsSource = null;
            DGMediosDePagos.Items.Clear();
            cheques.Clear();
            textBlock3.Text = "";
            textBlock5.Text = "";
            txtVuelto.Text = "";
            btnConfirPag.IsEnabled = false;
            RFC_Combo_Bancos();
            GC.Collect();
        }

        public void LimpiarPagMasivo()
        {
            DGChequeMasivo.ItemsSource = null;
            DGChequeMasivo.Items.Clear();
            DGMediosDePagosMasivo.ItemsSource = null;
            DGMediosDePagosMasivo.Items.Clear();
            chequesMasiv.Clear();
            btnAgregaMtoMasivo.Visibility = Visibility.Visible;
            btnBuscarPM.Visibility = Visibility.Collapsed;
            PrgBarExcel.Visibility = Visibility.Collapsed;
            textBlock3Masivo.Text = "";
            txtMontoFPMasivo.Text = textBlock4Masivo.Text;
            textBlock5Masivo.Text = "";
            btnConfirPag.IsEnabled = false;
            RFC_Combo_Bancos();
            txtVuelto.Visibility = Visibility.Collapsed;
            lbvuelto.Visibility = Visibility.Collapsed;
            GC.Collect();
        }

        public void LimpiarPagosMasivos(object sender, RoutedEventArgs e)
        {
            DGChequeMasivo.ItemsSource = null;
            DGChequeMasivo.Items.Clear();
            DGMediosDePagosMasivo.ItemsSource = null;
            DGMediosDePagosMasivo.Items.Clear();
            cmbVPMedioPagMasivo.ItemsSource = null;
            cmbVPMedioPagMasivo.Items.Clear();
            btnAgregaMtoMasivo.Visibility = Visibility.Visible;
            btnBuscarPM.Visibility = Visibility.Collapsed;
            PrgBarExcel.Visibility = Visibility.Collapsed;
            chequesMasiv.Clear();
            textBlock3Masivo.Text = "";
            txtMontoFPMasivo.Text = textBlock4Masivo.Text;
            textBlock5Masivo.Text = "";
            btnConfirPag.IsEnabled = false;
            RFC_Combo_Bancos();
            GC.Collect();
        }

        //BOTON QUE LLEVA EL DATO SELECCIONADO DESDE EL MONITOR 
        private void btnPagoMonitor_Click(object sender, RoutedEventArgs e)
        {
            LimpiarViasDePago();
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            ListaDocumentosPendientesDesdeMonitor();
            GBDocsAPagar.Visibility = Visibility.Visible;
            GBViasPago.Visibility = Visibility.Visible;
            txtDocu.Text = "";
            txtDocuAnt.Text = "";
            txtRut.Text = "";
            txtRUTAnt.Text = "";
            GC.Collect();
        }
        #endregion
        //MANEJO DE LOS EVENTOS ASOCIADOS A LOS RADIOBUTTONS
        #region RadioButtons
        private void RBRUTAnt_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTAnt.Text = "";
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            txtRUTAnt.Visibility = Visibility.Visible;
            txtDocuAnt.Text = "";
            txtDocuAnt.Visibility = Visibility.Collapsed;
            anticipos = new Anticipos();
            btnBuscarAnt.IsEnabled = true;
            LimpiarViasDePago();
            GC.Collect();
        }

        private void RBDocuAnt_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTAnt.Text = "";
            txtRUTAnt.Visibility = Visibility.Collapsed;
            txtDocuAnt.Text = "";
            txtDocuAnt.Visibility = Visibility.Visible;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            anticipos = new Anticipos();
            btnBuscarAnt.IsEnabled = true;
            LimpiarViasDePago();
            GC.Collect();
        }

        private void RBRut_Checked(object sender, RoutedEventArgs e)
        {
            txtRut.Text = "";
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            txtRut.Visibility = Visibility.Visible;
            txtDocu.Text = "";
            txtDocu.Visibility = Visibility.Collapsed;
            partidasabiertas = new PartidasAbiertas();
            btnBuscarP.IsEnabled = true;
            LimpiarViasDePago();
            GC.Collect();
        }

        private void RBDoc_Checked(object sender, RoutedEventArgs e)
        {
            txtRut.Text = "";
            txtRut.Visibility = Visibility.Collapsed;
            txtDocu.Text = "";
            txtDocu.Visibility = Visibility.Visible;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            partidasabiertas = new PartidasAbiertas();
            btnBuscarP.IsEnabled = true;
            LimpiarViasDePago();
            GC.Collect();
        }

        private void RBRutRE_Checked(object sender, RoutedEventArgs e)
        {
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Visible;
            btnBuscarP.IsEnabled = true;
            LimpiarViasDePago();
            GC.Collect();
        }

        private void RBDocRE_Checked(object sender, RoutedEventArgs e)
        {

            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            btnBuscarP.IsEnabled = true;
            LimpiarViasDePago();
            GC.Collect();
        }

        #endregion

        //MANEJO DE LOS EVENTOS ASOCIADOS A SELECCION DE LINEAS EN DATAGRID
        #region // SelectionChanged
        private void DGMonitor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            timer.Stop();
            GC.Collect();
        }
        #endregion
        //MANEJO DE LOS EVENTOS ASOCIADOS A LOS COMBOBOXS
        #region // ComboBox
        //RFC QUE LLENA LAS TARJETAS EN MEDIOS DE PAGOS
        private void RFC_Combo_Tarjetas()
        {
            cmbTipoTarjeta.ItemsSource = null;
            cmbTipoTarjeta.Items.Clear();
            cmbTipoTarjetaMasivo.ItemsSource = null;
            cmbTipoTarjetaMasivo.Items.Clear();

            if (cmbVPMedioPag.Text == "")
            {
                maestrotarjetas.maestrotarjetas(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbVPMedioPagMasivo.Text));
            }
            else
            {
                maestrotarjetas.maestrotarjetas(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbVPMedioPag.Text));
            }
            if (maestrotarjetas.T_Retorno.Count > 0)
            {
                cmbTipoTarjeta.ItemsSource = null;
                cmbTipoTarjeta.Items.Clear();

                cmbTipoTarjetaMasivo.ItemsSource = null;
                cmbTipoTarjetaMasivo.Items.Clear();

                List<string> listatarjetas = new List<string>();
                listatarjetas.Clear();
                for (int i = 0; i < maestrotarjetas.T_Retorno.Count; i++)
                {
                    listatarjetas.Add(maestrotarjetas.T_Retorno[i].CCINS + " - " + maestrotarjetas.T_Retorno[i].VTEXT);
                }
                cmbTipoTarjeta.ItemsSource = listatarjetas;
                cmbTipoTarjetaMasivo.ItemsSource = listatarjetas;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos de " + Convert.ToString(cmbVPMedioPag.Text).Substring(3, Convert.ToString(cmbVPMedioPag.Text).Length - 3) + " en el sistema");
            }
            GC.Collect();
        }
        private void RFC_Carta_Curse()
        {
            maestrofinanc.maestroifinan(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), "K", Convert.ToString(lblSociedad.Content));
            if (maestrofinanc.T_Retorno.Count > 0)
            {

                cmbIfinan.ItemsSource = null;
                cmbIfinan.Items.Clear();
                cmbIfinanMasivo.ItemsSource = null;
                cmbIfinanMasivo.Items.Clear();

                List<string> listabancos = new List<string>();
                listabancos.Clear();
                for (int i = 0; i < maestrofinanc.T_Retorno.Count; i++)
                {
                    listabancos.Add(maestrofinanc.T_Retorno[i].KUNNR + " - " + maestrofinanc.T_Retorno[i].MCOD1);
                }
                cmbIfinan.ItemsSource = listabancos;
                cmbIfinanMasivo.ItemsSource = listabancos;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos de instituciones financieras en el sistema");
            }
            GC.Collect();
        }

        //RFC QUE LLENA LOS BANCOS EN MEDIOS DE PAGOS

        private void RFC_Combo_Bancos()
        {
            maestrobancos.maestrobancos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbMoneda.Text), Convert.ToString(lblSociedad.Content));
            if (maestrobancos.T_Retorno.Count > 0)
            {
                cmbBancoProp.ItemsSource = null;
                cmbBancoProp.Items.Clear();
                cmbCuentasBancosProp.ItemsSource = null;
                cmbCuentasBancosProp.Items.Clear();
                cmbBanco.ItemsSource = null;
                cmbBanco.Items.Clear();
                 
                cmbBancoPropMasivoMasivo.ItemsSource = null;
                cmbBancoPropMasivoMasivo.Items.Clear();
                cmbCuentasBancosPropMasivo.ItemsSource = null;
                cmbCuentasBancosPropMasivo.Items.Clear();
                cmbBancoMasivo.ItemsSource = null;
                cmbBancoMasivo.Items.Clear();


                List<string> listabancos = new List<string>();
                listabancos.Clear();

                for (int i = 0; i < maestrobancos.T_Retorno.Count; i++)
                {
                    listabancos.Add(maestrobancos.T_Retorno[i].BANKL + " - " + maestrobancos.T_Retorno[i].BANKA);
                }
                cmbBanco.ItemsSource = listabancos;
                cmbBancoMasivo.ItemsSource = listabancos;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos de bancos en el sistema");
            }
            GC.Collect();
        }
        //RFC PARA EL USO DE BANCOS PROPIOS Y CUENTAS ASOCIADAS EN MEDIOS DE PAGOS
        private void RFC_Combo_BancosPropios()
        {
            string ViaPago = string.Empty;
            string[] split = cmbVPMedioPag.SelectedValue.ToString().Split(new Char[] { '-' });
            ViaPago = split[0];

            if (ViaPago == "O ")
            {
                maestrobancos.maestrobancos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(""), Convert.ToString(lblSociedad.Content));
            }
            else
            {
                maestrobancos.maestrobancos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbMoneda.Text), Convert.ToString(lblSociedad.Content));
            }
            if (maestrobancos.T_Retorno2.Count > 0)
            {
                cmbBancoProp.ItemsSource = null;
                cmbBancoProp.Items.Clear();
                cmbCuentasBancosProp.ItemsSource = null;
                cmbCuentasBancosProp.Items.Clear();
                cmbBanco.ItemsSource = null;
                cmbBanco.Items.Clear();

                // Limpia Combobox Bancos Propios Pagos Masivos
                cmbBancoPropMasivoMasivo.ItemsSource = null;
                cmbBancoPropMasivoMasivo.Items.Clear();
                cmbCuentasBancosPropMasivo.ItemsSource = null;
                cmbCuentasBancosPropMasivo.Items.Clear();
                cmbBancoMasivo.ItemsSource = null;
                cmbBancoMasivo.Items.Clear();

                List<string> listabancos = new List<string>();
                List<string> listabancosprop = new List<string>();
                List<string> cuentasbancosprop = new List<string>();

                listabancos.Clear();
                listabancosprop.Clear();
                cuentasbancosprop.Clear();

                for (int i = 0; i < maestrobancos.T_Retorno.Count; i++)
                {
                    if (!listabancos.Contains(maestrobancos.T_Retorno[i].BANKL + " - " + maestrobancos.T_Retorno[i].BANKA))
                    {
                        listabancos.Add(maestrobancos.T_Retorno[i].BANKL + " - " + maestrobancos.T_Retorno[i].BANKA);
                    }
                }

                for (int i = 0; i < maestrobancos.T_Retorno2.Count; i++)
                {
                    if (!listabancosprop.Contains(maestrobancos.T_Retorno2[i].HBKID))
                    {
                        listabancosprop.Add(maestrobancos.T_Retorno2[i].HBKID);
                    }
                    if (!cuentasbancosprop.Contains(maestrobancos.T_Retorno2[i].BANKN))
                    {
                        cuentasbancosprop.Add(maestrobancos.T_Retorno2[i].BANKN);
                    }
                }
                cmbBanco.ItemsSource = listabancos;
                cmbBancoProp.ItemsSource = listabancosprop;
                cmbCuentasBancosProp.ItemsSource = cuentasbancosprop;

                cmbBancoMasivo.ItemsSource = listabancos;
                cmbBancoPropMasivoMasivo.ItemsSource = listabancosprop;
                cmbCuentasBancosPropMasivo.ItemsSource = cuentasbancosprop;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos de bancos propios en el sistema");
            }
            GC.Collect();
        }

        private void cmbBancoPropMasivo_DropDownClosed(object sender, EventArgs e)
        {
            int posicion;

            posicion = cmbBancoPropMasivoMasivo.SelectedIndex;
            cmbCuentasBancosPropMasivo.SelectedIndex = posicion;
            GC.Collect();
        }
        //SELECCION DEL COMBOBOX DE BANCOS PROPIOS
        private void cmbBancoProp_DropDownClosed(object sender, EventArgs e)
        {
            int posicion;
            posicion = cmbBancoProp.SelectedIndex;
            cmbCuentasBancosProp.SelectedIndex = posicion;
            GC.Collect();
        }
        //LLena Medios De Pagos Masivos
        private void cmbVPMedioPagMasivo_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                if (cmbVPMedioPagMasivo.Text != "")
                {
                    string MedioPago = cmbVPMedioPagMasivo.Text.Substring(0, 1);
                    cmbBancoPropMasivoMasivo.ItemsSource = null;
                    cmbBancoPropMasivoMasivo.Items.Clear();
                    cmbCuentasBancosPropMasivo.ItemsSource = null;
                    cmbCuentasBancosPropMasivo.Items.Clear();
                    cmbBancoMasivo.ItemsSource = null;
                    cmbBancoMasivo.Items.Clear();
                    switch (MedioPago)
                    {
                        case "K": //Carta curse
                            {
                                RFC_Combo_Bancos();
                                RFC_Carta_Curse();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Visible;
                                cmbIfinanMasivo.Visibility = Visibility.Visible;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "F": //Cheque a fecha
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Visible;
                                txtCodAutMasivo.Visibility = Visibility.Visible;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Visible;
                                cmbBancoMasivo.Visibility = Visibility.Visible;
                                label27Masivo.Visibility = Visibility.Visible;
                                txtSucursalMasivo.Visibility = Visibility.Visible;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Visible;
                                txtNumCuentaMasivo.Visibility = Visibility.Visible;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Visible;
                                label33Masivo.Content = "N° de documentos";
                                txtCantDocMasivo.Visibility = Visibility.Visible;
                                label34Masivo.Visibility = Visibility.Visible;
                                cmbIntervaloMasivo.Visibility = Visibility.Visible;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "G": //Cheque al día
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Visible;
                                txtCodAutMasivo.Visibility = Visibility.Visible;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Visible;
                                cmbBancoMasivo.Visibility = Visibility.Visible;
                                label27Masivo.Visibility = Visibility.Visible;
                                txtSucursalMasivo.Visibility = Visibility.Visible;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Visible;
                                txtNumCuentaMasivo.Visibility = Visibility.Visible;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "M": //Contrato compra-venta
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Visible;
                                txtObservMasivo.Visibility = Visibility.Visible;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "D": //Deposito a plazo
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Visible;
                                cmbBancoMasivo.Visibility = Visibility.Visible;
                                label27Masivo.Visibility = Visibility.Visible;
                                txtSucursalMasivo.Visibility = Visibility.Visible;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "B": //Deposito en cuenta corriente
                            {
                                //RFC BANCO PROPIO
                                RFC_Combo_BancosPropios();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Visible;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Visible;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Visible;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Visible;
                                txtObservMasivo.Visibility = Visibility.Visible;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                break;
                            }
                        case "L": //Letras
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Visible;
                                txtCantDocMasivo.Visibility = Visibility.Visible;
                                label34Masivo.Visibility = Visibility.Visible;
                                cmbIntervaloMasivo.Visibility = Visibility.Visible;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "P": //Pagaré
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Visible;
                                txtCantDocMasivo.Visibility = Visibility.Visible;
                                label34Masivo.Visibility = Visibility.Visible;
                                cmbIntervaloMasivo.Visibility = Visibility.Visible;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "E": //Pago en efectivo
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Collapsed;
                                DPFechActualMasivo.Visibility = Visibility.Collapsed;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Collapsed;
                                txtNumDocMasivo.Visibility = Visibility.Collapsed;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Visible;
                                txtObservMasivo.Visibility = Visibility.Visible;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "H": //Pago Dolares
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Collapsed;
                                DPFechActualMasivo.Visibility = Visibility.Collapsed;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Collapsed;
                                txtNumDocMasivo.Visibility = Visibility.Collapsed;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Visible;
                                txtObservMasivo.Visibility = Visibility.Visible;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "S": //Tarjeta de crédito 
                            {

                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Visible;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Visible;
                                label48Masivo.Visibility = Visibility.Visible;
                                txtCodAutMasivo.Visibility = Visibility.Visible;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Visible;
                                txtCodOpMasivo.Visibility = Visibility.Visible;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Collapsed;
                                DPFechActualMasivo.Visibility = Visibility.Collapsed;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Visible;
                                label33Masivo.Content = "Número de cuotas";
                                txtCantDocMasivo.Visibility = Visibility.Visible;
                                label34Masivo.Visibility = Visibility.Visible;
                                cmbIntervaloMasivo.Visibility = Visibility.Visible;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "R": //Tarjeta de débito
                            {
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Visible;
                                txtCodAutMasivo.Visibility = Visibility.Visible;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Visible;
                                txtCodOpMasivo.Visibility = Visibility.Visible;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Collapsed;
                                DPFechActualMasivo.Visibility = Visibility.Collapsed;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }

                        case "N": // Servipag
                            {
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Visible;
                                txtCodAutMasivo.Visibility = Visibility.Visible;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Visible;
                                txtCodOpMasivo.Visibility = Visibility.Visible;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Visible;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Collapsed;
                                DPFechActualMasivo.Visibility = Visibility.Collapsed;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Collapsed;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }

                        case "U": //Transferencia bancaria
                            {
                                //RFC BANCO PROPIO
                                RFC_Combo_BancosPropios();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Visible;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Visible;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Visible;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "V": //Vale vista recibido
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Visible;
                                DPFechVencMasivo.Visibility = Visibility.Visible;
                                label26Masivo.Visibility = Visibility.Visible;
                                cmbBancoMasivo.Visibility = Visibility.Visible;
                                label27Masivo.Visibility = Visibility.Visible;
                                txtSucursalMasivo.Visibility = Visibility.Visible;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Content = "Número venta";
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "A": //Vehiculo en parte de pago
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Visible;
                                DPFechActualMasivo.Visibility = Visibility.Visible;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Visible;
                                txtNumDocMasivo.Visibility = Visibility.Visible;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Visible;
                                txtRUTGiradorMasivo.Visibility = Visibility.Visible;
                                label38Masivo.Visibility = Visibility.Visible;
                                txtNombreGiraMasivo.Visibility = Visibility.Visible;
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Visible;
                                txtPatenteMasivo.Visibility = Visibility.Visible;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        case "1": //Documento Tributario
                            {
                                RFC_Combo_Bancos();
                                label46Masivo.Visibility = Visibility.Collapsed;
                                cmbTipoTarjetaMasivo.Visibility = Visibility.Collapsed;
                                label48Masivo.Visibility = Visibility.Collapsed;
                                txtCodAutMasivo.Visibility = Visibility.Collapsed;
                                btnAutorizacionMasivo.Visibility = Visibility.Collapsed;
                                label49Masivo.Visibility = Visibility.Collapsed;
                                txtCodOpMasivo.Visibility = Visibility.Collapsed;
                                label50Masivo.Visibility = Visibility.Collapsed;
                                txtAsigMasivo.Visibility = Visibility.Collapsed;
                                label43Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoPropMasivoMasivo.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosPropMasivo.Visibility = Visibility.Collapsed;
                                label32Masivo.Visibility = Visibility.Collapsed;
                                DPFechActualMasivo.Visibility = Visibility.Collapsed;
                                label25Masivo.Visibility = Visibility.Collapsed;
                                DPFechVencMasivo.Visibility = Visibility.Collapsed;
                                label26Masivo.Visibility = Visibility.Collapsed;
                                cmbBancoMasivo.Visibility = Visibility.Collapsed;
                                label27Masivo.Visibility = Visibility.Collapsed;
                                txtSucursalMasivo.Visibility = Visibility.Collapsed;
                                label28Masivo.Visibility = Visibility.Collapsed;
                                txtNumDocMasivo.Visibility = Visibility.Collapsed;
                                label30Masivo.Visibility = Visibility.Collapsed;
                                txtNumCuentaMasivo.Visibility = Visibility.Collapsed;
                                label31Masivo.Visibility = Visibility.Collapsed;
                                txtRUTGiradorMasivo.Visibility = Visibility.Collapsed;
                                label38Masivo.Visibility = Visibility.Collapsed;
                                txtNombreGiraMasivo.Visibility = Visibility.Collapsed;
                                label39Masivo.Visibility = Visibility.Collapsed;
                                txtNumVentaMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.Visibility = Visibility.Collapsed;
                                txtObservMasivo.Visibility = Visibility.Collapsed;
                                label42Masivo.Visibility = Visibility.Collapsed;
                                cmbIfinanMasivo.Visibility = Visibility.Collapsed;
                                lblPatenteMasivo.Visibility = Visibility.Collapsed;
                                txtPatenteMasivo.Visibility = Visibility.Collapsed;
                                label34Masivo.Visibility = Visibility.Collapsed;
                                txtCantDocMasivo.Visibility = Visibility.Collapsed;
                                label33Masivo.Visibility = Visibility.Collapsed;
                                cmbIntervaloMasivo.Visibility = Visibility.Collapsed;
                                label40Masivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40Masivo.Margin = new Thickness(16, 64, 0, 0);
                                txtObservMasivo.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObservMasivo.Margin = new Thickness(136, 64, 0, 0);
                                txtObservMasivo.Width = 389;
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
            GC.Collect();

        }
        //MUESTRA (VISIBILIDAD DE LOS ITEMS DE LOS MEDIOS DE PAGO A PARTIR DE LA SELECCION DEL COMBO DE MEDIOS DE PAGO
        private void cmbVPMedioPag_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                LimpiarEntradaDeViasDePago();
                if (cmbVPMedioPag.Text != "")
                {
                    string MedioPago = cmbVPMedioPag.Text.Substring(0, 1);
                    cmbBancoProp.ItemsSource = null;
                    cmbBancoProp.Items.Clear();
                    cmbCuentasBancosProp.ItemsSource = null;
                    cmbCuentasBancosProp.Items.Clear();
                    cmbBanco.ItemsSource = null;
                    cmbBanco.Items.Clear();
                    switch (MedioPago)
                    {
                        case "K": //Carta curse
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                RFC_Carta_Curse();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Visible;
                                cmbIfinan.Visibility = Visibility.Visible;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "F": //Cheque a fecha
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Visible;
                                txtCodAut.Visibility = Visibility.Visible;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Visible;
                                cmbBanco.Visibility = Visibility.Visible;
                                label27.Visibility = Visibility.Visible;
                                txtSucursal.Visibility = Visibility.Visible;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Visible;
                                txtNumCuenta.Visibility = Visibility.Visible;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Visible;
                                label33.Content = "N° de documentos";
                                txtCantDoc.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Visible;
                                cmbIntervalo.Visibility = Visibility.Visible;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "G": //Cheque al día
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Visible;
                                txtCodAut.Visibility = Visibility.Visible;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Visible;
                                cmbBanco.Visibility = Visibility.Visible;
                                label27.Visibility = Visibility.Visible;
                                txtSucursal.Visibility = Visibility.Visible;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Visible;
                                txtNumCuenta.Visibility = Visibility.Visible;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "M": //Contrato compra-venta
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Visible;
                                txtObserv.Visibility = Visibility.Visible;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "D": //Deposito a plazo
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Visible;
                                cmbBanco.Visibility = Visibility.Visible;
                                label27.Visibility = Visibility.Visible;
                                txtSucursal.Visibility = Visibility.Visible;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "B": //Deposito en cliente corriente
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_BancosPropios();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Visible;
                                cmbBancoProp.Visibility = Visibility.Visible;
                                cmbCuentasBancosProp.Visibility = Visibility.Visible;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Visible;
                                txtObserv.Visibility = Visibility.Visible;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                break;
                            }
                        case "L": //Letras
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Visible;
                                txtCantDoc.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Visible;
                                cmbIntervalo.Visibility = Visibility.Visible;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "P": //Pagaré
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Visible;
                                txtCantDoc.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Visible;
                                cmbIntervalo.Visibility = Visibility.Visible;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "J": //Pago en Euros.
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    string convert2 = convert.Replace(".", "");
                                    int diferen = Convert.ToInt32(convert2);

                                    if (diferen >= 0)
                                    {
                                        btnValidarEfect.IsEnabled = true;
                                        convert = textBlock5.Text;
                                        convert2 = convert.Replace(".", "");
                                        diferen = Convert.ToInt32(convert2);
                                        if (diferen > 0)
                                        {
                                            string resul2 = Conversion(textBlock5.Text);
                                            txtMontoFP.Text = resul2;
                                            txt_dolar.Text = resul2;
                                        }
                                        else
                                        {
                                            txtMontoFP.Text = "";
                                            string resul = Conversion(textBlock4.Text);
                                            txtMontoFP.Text = resul;
                                            txt_dolar.Text = resul;
                                        }
                                    }
                                }
                                else
                                {
                                    txtMontoFP.Text = "";
                                    string resul = Conversion(textBlock4.Text);
                                    txtMontoFP.Text = resul;
                                    txt_dolar.Text = resul;
                                }
                                txtVuelto.Visibility = Visibility.Visible;
                                lbvuelto.Visibility = Visibility.Visible;
                                RFC_Combo_Bancos();
                                btn_conver.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Collapsed;
                                txtNumDoc.Visibility = Visibility.Collapsed;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Visible;
                                txtObserv.Visibility = Visibility.Visible;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "H": //Pago en dolares
                            {   
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    string convert2 = convert.Replace(".", "");
                                    int diferen = Convert.ToInt32(convert2);

                                    if (diferen >= 0)
                                    {
                                        btnValidarEfect.IsEnabled = true;
                                        convert = textBlock5.Text;
                                        convert2 = convert.Replace(".", "");
                                        diferen = Convert.ToInt32(convert2);
                                        if (diferen > 0)
                                        {
                                            string resul2 = Conversion(textBlock5.Text);
                                            txtMontoFP.Text = resul2;
                                            txt_dolar.Text = resul2;
                                        }
                                        else
                                        {
                                            txtMontoFP.Text = "";
                                            string resul = Conversion(textBlock4.Text);
                                            txtMontoFP.Text = resul;
                                            txt_dolar.Text = resul;
                                        }
                                    }
                                }
                                else
                                {
                                    txtMontoFP.Text = "";
                                    string resul = Conversion(textBlock4.Text);
                                    txtMontoFP.Text = resul;
                                    txt_dolar.Text = resul;
                                }
                                txtVuelto.Visibility = Visibility.Collapsed;
                                lbvuelto.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                btn_conver.IsEnabled = true;
                                lbDolar.Visibility = Visibility.Collapsed;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Collapsed;
                                txtNumDoc.Visibility = Visibility.Collapsed;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Visible;
                                txtObserv.Visibility = Visibility.Visible;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "E": //Pago en efectivo
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                btn_conver.Visibility = Visibility.Collapsed;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Collapsed;
                                txtNumDoc.Visibility = Visibility.Collapsed;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Visible;
                                txtObserv.Visibility = Visibility.Visible;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "S": //Tarjeta de crédito 
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Visible;
                                cmbTipoTarjeta.Visibility = Visibility.Visible;
                                label48.Visibility = Visibility.Visible;
                                txtCodAut.Visibility = Visibility.Visible;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Visible;
                                txtCodOp.Visibility = Visibility.Visible;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Visible;
                                label33.Content = "Número de cuotas";
                                txtCantDoc.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Visible;
                                cmbIntervalo.Visibility = Visibility.Visible;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "R": //Tarjeta de débito
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Visible;
                                txtCodAut.Visibility = Visibility.Visible;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Visible;
                                txtCodOp.Visibility = Visibility.Visible;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "I": //GIFTCARD
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Visible;
                                txtCodAut.Visibility = Visibility.Visible;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Visible;
                                txtCodOp.Visibility = Visibility.Visible;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "N": //Servipag
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "U": //Transferencia bancaria
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                //RFC BANCO PROPIO
                                RFC_Combo_BancosPropios();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Visible;
                                cmbBancoProp.Visibility = Visibility.Visible;
                                cmbCuentasBancosProp.Visibility = Visibility.Visible;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "O": //Transferencia bancaria Dolares
                            {
                                string convert = textBlock4.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    string convert2 = convert.Replace(".", "");
                                    int diferen = Convert.ToInt32(convert2);

                                    if (diferen >= 0)
                                    {
                                        btnValidarEfect.IsEnabled = true;
                                        convert = textBlock4.Text;
                                        convert2 = convert.Replace(".", "");
                                        diferen = Convert.ToInt32(convert2);
                                        if (diferen > 0)
                                        {
                                            string resul2 = Conversion(textBlock4.Text);
                                            txtMontoFP.Text = resul2;
                                            txt_dolar.Text = resul2;
                                        }
                                        else
                                        {
                                            txtMontoFP.Text = "";
                                            string resul = Conversion(textBlock5.Text);
                                            txtMontoFP.Text = resul;
                                            txt_dolar.Text = resul;
                                        }
                                    }
                                }
                                else
                                {
                                    txtMontoFP.Text = "";
                                    string resul = Conversion(textBlock5.Text);
                                    txtMontoFP.Text = resul;
                                    txt_dolar.Text = resul;
                                }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                //RFC BANCO PROPIO
                                RFC_Combo_BancosPropios();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Visible;
                                cmbBancoProp.Visibility = Visibility.Visible;
                                cmbCuentasBancosProp.Visibility = Visibility.Visible;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "V": //Vale vista recibido
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;
                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Visible;
                                DPFechVenc.Visibility = Visibility.Visible;
                                label26.Visibility = Visibility.Visible;
                                cmbBanco.Visibility = Visibility.Visible;
                                label27.Visibility = Visibility.Visible;
                                txtSucursal.Visibility = Visibility.Visible;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Content = "Número venta";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "A": //Vehiculo en parte de pago
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;

                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Visible;
                                DPFechActual.Visibility = Visibility.Visible;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Visible;
                                txtNumDoc.Visibility = Visibility.Visible;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Visible;
                                txtRUTGirador.Visibility = Visibility.Visible;
                                label38.Visibility = Visibility.Visible;
                                txtNombreGira.Visibility = Visibility.Visible;
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Visible;
                                txtPatente.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;

                            }
                        case "1": //Documento Tributario
                            {
                                string convert = textBlock5.Text;
                                if (!String.IsNullOrEmpty(convert))
                                {
                                    txtMontoFP.Text = convert;
                                }
                                else { txtMontoFP.Text = textBlock4.Text; }
                                btnValidarEfect.IsEnabled = false;
                                txt_dolar.Visibility = Visibility.Collapsed;
                                txt_euros.Visibility = Visibility.Collapsed;
                                lbDolar.Visibility = Visibility.Collapsed;
                                lbEuro.Visibility = Visibility.Collapsed;
                                btn_conver.Visibility = Visibility.Collapsed;

                                RFC_Combo_Bancos();
                                txtMontoFP.IsEnabled = true;
                                label46.Visibility = Visibility.Collapsed;
                                cmbTipoTarjeta.Visibility = Visibility.Collapsed;
                                label48.Visibility = Visibility.Collapsed;
                                txtCodAut.Visibility = Visibility.Collapsed;
                                btnAutorizacion.Visibility = Visibility.Collapsed;
                                label49.Visibility = Visibility.Collapsed;
                                txtCodOp.Visibility = Visibility.Collapsed;
                                label50.Visibility = Visibility.Collapsed;
                                txtAsig.Visibility = Visibility.Collapsed;
                                label43.Visibility = Visibility.Collapsed;
                                cmbBancoProp.Visibility = Visibility.Collapsed;
                                cmbCuentasBancosProp.Visibility = Visibility.Collapsed;
                                label32.Visibility = Visibility.Collapsed;
                                DPFechActual.Visibility = Visibility.Collapsed;
                                label25.Visibility = Visibility.Collapsed;
                                DPFechVenc.Visibility = Visibility.Collapsed;
                                label26.Visibility = Visibility.Collapsed;
                                cmbBanco.Visibility = Visibility.Collapsed;
                                label27.Visibility = Visibility.Collapsed;
                                txtSucursal.Visibility = Visibility.Collapsed;
                                label28.Visibility = Visibility.Collapsed;
                                txtNumDoc.Visibility = Visibility.Collapsed;
                                label30.Visibility = Visibility.Collapsed;
                                txtNumCuenta.Visibility = Visibility.Collapsed;
                                label31.Visibility = Visibility.Collapsed;
                                txtRUTGirador.Visibility = Visibility.Collapsed;
                                label38.Visibility = Visibility.Collapsed;
                                txtNombreGira.Visibility = Visibility.Collapsed;
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                lb_codauto.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
            GC.Collect();
        }
        #endregion
        //FUNCIONES y METODOS
        #region // Funciones
        //FUNCION QUE CONTROLA LA LECTURA DE DATOS DE EL MONITOR
        void timer_Tick(object sender, EventArgs e)
        {
            try
            {
                monitor.ObjDatosMonitor.Clear();
                monitor.monitor(Convert.ToString(Calendario.SelectedDate.Value), Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblSociedad.Content));
                if (monitor.ObjDatosMonitor.Count > 0)
                {
                    DGMonitor.ItemsSource = null;
                    DGMonitor.Items.Clear();
                    DGMonitor.ItemsSource = monitor.ObjDatosMonitor;
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        //FUNCION QUE HACE LA REVISION DEL DIGITO VERIFICADOR
        private string DigitoVerificador(string RUTU)
        {
            string digito = "";
            string RUTsDV = "";
            string RUTcDV = "";
            string RUT = "";
            int Total = 0;
            int j = 2;
            int modulo = 0;
            foreach (char value in RUTU)
            {
                if (value != 8207)
                {
                    RUT = RUT + value;
                }
            }

            RUT = RUT.Trim();
            if (RUT.Contains("-"))
            {
                RUT = RUT.Substring(0, RUT.Length - 2);
            }

            for (int i = RUT.Length - 1; i >= 0; --i)
            {
                RUTsDV = RUTsDV + RUT[i];
            }
            for (int i = 0; i < RUTsDV.Length; i++)
            {
                digito = Convert.ToString(RUTsDV[i]);
                Total = Total + Convert.ToInt16(digito) * j;
                if (j < 7)
                {
                    j++;
                }
                else
                {
                    j = 2;
                }
            }
            modulo = Total % 11;
            Total = 11 - modulo;
            digito = "";
            switch (Total)
            {
                case 11: //Apertura de caja exitosa
                    {
                        digito = "0";
                        break;
                    }
                case 10: //Acceso a caja por usuario distinto al que realizo la apertura
                    {
                        digito = "K";
                        break;
                    }

                default:
                    {
                        digito = Convert.ToString(Total);
                        break;
                    }
            }
            RUTcDV = RUT + "-" + digito;
            return RUTcDV;
            GC.Collect();
        }
        //FUNCION QUE TOMA EL INGRESO DE DETALLES DE MEDIOS Y FORMAS DE PAGO, MONTOS Y LLENA EL GRID RESUMEN DE PAGOS (DGCHEQUES).
        private void IngresoFormasDePagoYMontos(string MedioPago)
        {
            try
            {


                String FechaVenct;
                string Valor = string.Empty;

                if (DPFechActual.Text == "")
                {
                    DPFechActual.Text = Calendario.Text;
                    DPFechActual.Text = Convert.ToString(Calendario.SelectedDate.Value);
                }

                if (DPFechVenc.Text != "")
                {
                    FechaVenct = DPFechVenc.Text.Substring(0, 10);
                }
                else
                {
                    FechaVenct = DPFechActual.Text.Substring(0, 10);
                }
                string mandt = "";
                string land = Convert.ToString(lblPais.Content);
                string id_comprobante = "";
                string id_detalle = "";
                string via_pago = MedioPago;

                if (via_pago == "H" || via_pago == "J"|| via_pago == "O")
                {
                    string resul = string.Empty;
                    switch (via_pago)
                    {

                        case "O":
                            //resul = Conversion(txtMontoFP.Text);
                            if (!txtMontoFP.Text.Equals(txt_dolar.Text))
                            {
                                resul = converOtrasMonedas(txtMontoFP.Text);

                                Valor = resul.Replace(".", "");
                                Valor = Valor.Replace(",", ".");
                                monto = Convert.ToDouble(Valor);
                            }
                            else
                            {
                                resul = txtMontoFP.Text;
                                resul = converOtrasMonedas(txtMontoFP.Text);

                                Valor = resul.Replace(".", "");
                                Valor = Valor.Replace(",", ".");
                                monto = Math.Ceiling(Convert.ToDouble(Valor));

                            }
                            //Eliminación de la separación de miles

                            Valor2 = txtMontoFP.Text.Replace(".", "");
                            Valor2 = Valor2.Replace(",", ".");
                            monto2 = Convert.ToDouble(Valor2);
                            break;

                        case "H":
                            //resul = Conversion(txtMontoFP.Text);
                            if (!txtMontoFP.Text.Equals(txt_dolar.Text))
                            {
                                resul = converOtrasMonedas(txtMontoFP.Text);

                                Valor = resul.Replace(".", "");
                                Valor = Valor.Replace(",", ".");
                                monto = Convert.ToDouble(Valor);
                            }
                            else
                            {
                                resul = txtMontoFP.Text;
                                resul = converOtrasMonedas(txtMontoFP.Text);

                                Valor = resul.Replace(".", "");
                                Valor = Valor.Replace(",", ".");
                                monto = Math.Ceiling(Convert.ToDouble(Valor));
                                
                            }
                            //Eliminación de la separación de miles

                            Valor2 = txtMontoFP.Text.Replace(".", "");
                            Valor2 = Valor2.Replace(",", ".");
                            monto2 = Convert.ToDouble(Valor2);
                            break;
                        case "J":
                            //Eliminación de la separación de miles
                            if (!txtMontoFP.Text.Equals(txt_dolar.Text))
                            {
                                resul = converOtrasMonedas(txtMontoFP.Text);
                                Valor = resul.Replace(".", "");
                                Valor = Valor.Replace(",", ".");
                                monto = Convert.ToDouble(Valor);

                            }
                            else
                            {
                                resul = txtMontoFP.Text;
                                resul = converOtrasMonedas(txtMontoFP.Text);
                                Valor = resul.Replace(".", "");
                                Valor = Valor.Replace(",", ".");
                                monto = Convert.ToDouble(Valor);
                            }
                            //Eliminación de la separación de miles
                            Valor2 = txtMontoFP.Text.Replace(".", "");
                            Valor2 = Valor2.Replace(",", ".");
                            monto2 = Convert.ToDouble(Valor2);
                            break;
                    }
                }
                else
                {
                    //Eliminación de la separación de miles
                    Valor = txtMontoFP.Text.Replace(".", "");
                    Valor = Valor.Replace(",", ".");
                    monto3 = Convert.ToDouble(Valor);
                }

                if (via_pago == "H" | via_pago == "J" | via_pago == "O")
                {
                    switch (via_pago)
                    {
                        case "O":
                            moneda = "USD";

                            break;
                        case "H":
                            moneda = "USD";

                            break;
                        case "J":

                            moneda = "EUR";

                            break;
                    }
                }
                else
                {
                    moneda = cmbMoneda.Text;
                }

                string banco = cmbBanco.Text;
                string emisor = txtRUTGirador.Text;
                string num_cheque = "";
                string cod_autorizacion = txtCodAut.Text;
                int num_cuotas;
                int num_cuotasaux;
                if (txtCantDoc.Text == "")
                {
                    txtCantDoc.Text = "0";
                }
                num_cuotas = Convert.ToInt16(txtCantDoc.Text);

                string fecha_venc = FechaVenct;
                string texto_posicion = txtObserv.Text;
                string anexo = "";
                string sucursal = txtSucursal.Text;
                string num_cuenta = txtNumCuenta.Text;
                string num_tarjeta = "";
                string num_vale_vista = "";
                string patente = txtPatente.Text;
                string num_venta = txtNumVenta.Text;
                string pagare = "";
                string fecha_emision = DPFechActual.Text;
                string nombre_girador = txtNombreGira.Text;
                string carta_curse = "";
                string num_transfer = "";
                string num_deposito = "";
                string cta_banco = cmbCuentasBancosProp.Text;
                string ifinan = "";
                string corre = "";
                if (cmbIfinan.Text != "")
                {
                    int posicion = cmbIfinan.Text.LastIndexOf("-");
                    posicion = posicion - 1;
                    ifinan = cmbIfinan.Text.Substring(0, posicion);
                }
                else
                {
                    ifinan = "";
                }
                string zuonr = txtAsig.Text;
                string hkont = "";
                string prctr = "";
                string znop = txtCodOp.Text;
                string NumDoc = txtNumDoc.Text;
                string NumCtaCte = txtNumCuenta.Text;
                string Patente = txtPatente.Text;
                if (cmbIntervalo.Text == "")
                {
                    cmbIntervalo.Text = "0";
                }
                int Intervalo = Convert.ToInt16(cmbIntervalo.Text);

                DetalleViasPago detcheq;
                if (num_cuotas < 2) //VIAS DE PAGO SIN CUOTAS
                {
                    //DEPENDIENDO DEL MEDIO DE PAGO SE HACE LA LOGICA PARA INCORPORAR EL NUMERO DEL MEDIO DE PAGO AL CAMPO CORRECTO DE LA RFC
                    switch (MedioPago)
                    {
                        case "B":
                            {
                                num_deposito = txtNumDoc.Text;
                                break;
                            }
                        case "D":
                            {
                                num_deposito = txtNumDoc.Text;
                                break;
                            }
                        case "F":
                            {
                                num_cheque = txtNumDoc.Text;
                                break;
                            }
                        case "G":
                            {
                                num_cheque = txtNumDoc.Text;
                                break;
                            }
                        case "K":
                            {
                                carta_curse = txtNumDoc.Text;
                                num_venta = NumDoc;
                                break;
                            }
                        case "M":
                            {
                                pagare = txtNumDoc.Text;
                                break;
                            }
                        case "P":
                            {
                                pagare = txtNumDoc.Text;
                                break;
                            }
                        case "R":
                            {
                                num_tarjeta = txtNumDoc.Text;
                                break;
                            }
                        case "N":
                            {
                                num_tarjeta = txtNumDoc.Text;
                                break;
                            }
                        case "S":
                            {
                                num_tarjeta = txtNumDoc.Text;
                                break;
                            }
                        case "U":
                            {
                                num_transfer = txtNumDoc.Text;
                                break;
                            }
                        case "V":
                            {
                                num_vale_vista = txtNumDoc.Text;
                                break;
                            }
                        case "A":
                            {
                                num_venta = NumDoc;
                                break;
                            }
                    }
                    switch (via_pago)
                    {
                        case "H":
                            detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, monto2, moneda, banco, emisor
                            , num_cheque, cod_autorizacion, Convert.ToString(num_cuotas), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                            , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                            , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                            cheques.Add(detcheq);

                            break;
                        case "O":
                            detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, monto2, moneda, banco, emisor
                            , num_cheque, cod_autorizacion, Convert.ToString(num_cuotas), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                            , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                            , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                            cheques.Add(detcheq);

                            break;
                        case "J":
                            detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, monto2, moneda, banco, emisor
                                   , num_cheque, cod_autorizacion, Convert.ToString(num_cuotas), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                   , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                   , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                            cheques.Add(detcheq);

                            break;

                        default:
                            detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, monto3, moneda, banco, emisor
                           , num_cheque, cod_autorizacion, Convert.ToString(num_cuotas), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                           , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                           , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                            cheques.Add(detcheq);
                            break;
                    }
                }
                else
                {
                    double montotot = 0;
                    double montores = 0;

                    int j = 0;
                    if (monto3 % num_cuotas == 0) //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS EXACTO
                    {
                        for (int i = 1; i <= num_cuotas; i++)
                        {
                            double montoaux = Convert.ToDouble(monto3 / num_cuotas);
                            j++;

                            if (i != 1) //CALCULO DE LAS FECHAS DE VENCIMIENTO DE ACUERDO AL NUMERO DE CUOTAS
                            {
                                DateTime FechaVenctAux = Convert.ToDateTime(FechaVenct);
                                FechaVenctAux = FechaVenctAux.AddDays(Intervalo);
                                FechaVenct = Convert.ToString(FechaVenctAux);
                                fecha_venc = Convert.ToString(FechaVenctAux.ToString("dd/MM/yyyy"));
                                if (MedioPago != "S")
                                {
                                    NumDoc = Convert.ToString(Convert.ToInt64(NumDoc) + 1);
                                }
                            }
                            //DEPENDIENDO DEL MEDIO DE PAGO SE HACE LA LOGICA PARA INCORPORAR EL NUMERO DEL MEDIO DE PAGO AL CAMPO CORRECTO DE LA RFC
                            switch (MedioPago)
                            {
                                case "B":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "D":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "F":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "G":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "K":
                                    {
                                        num_venta = NumDoc;
                                        carta_curse = NumDoc;
                                        break;
                                    }
                                case "M":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "P":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "R":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "N":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "S":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "U":
                                    {
                                        num_transfer = NumDoc;
                                        break;
                                    }
                                case "V":
                                    {
                                        num_vale_vista = NumDoc;
                                        break;
                                    }
                                case "A":
                                    {
                                        num_venta = NumDoc;
                                        break;
                                    }
                            }
                            num_cuotasaux = 1;

                            detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, Convert.ToInt64(Math.Round(montoaux, 0)), moneda, banco, emisor
                                , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                            cheques.Add(detcheq);
                        }
                    }
                    else //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS CON UN RESIDUO QUE SE SUMA EN LA CUOTA FINAL
                    {
                        for (int i = 1; i <= num_cuotas; i++)
                        {
                            double montoaux = Convert.ToDouble(monto3 / num_cuotas);
                            montotot = montotot + Math.Round(montoaux, 0);
                            montores = (monto3) - montotot;
                            j++;
                            if (i != 1) //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS EXACTO
                            {
                                DateTime FechaVenctAux = Convert.ToDateTime(FechaVenct);
                                FechaVenctAux = FechaVenctAux.AddDays(Intervalo);
                                FechaVenct = Convert.ToString(FechaVenctAux);
                                fecha_venc = Convert.ToString(FechaVenctAux);
                                if (MedioPago != "S")
                                {
                                    NumDoc = Convert.ToString(Convert.ToInt64(NumDoc) + 1);
                                }
                            }
                            //DEPENDIENDO DEL MEDIO DE PAGO SE HACE LA LOGICA PARA INCORPORAR EL NUMERO DEL MEDIO DE PAGO AL CAMPO CORRECTO DE LA RFC
                            switch (MedioPago)
                            {
                                case "B":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "D":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "F":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "G":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "K":
                                    {
                                        carta_curse = NumDoc;
                                        num_venta = NumDoc;
                                        break;
                                    }
                                case "M":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "P":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "R":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "N":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "S":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "U":
                                    {
                                        num_transfer = NumDoc;
                                        break;
                                    }
                                case "V":
                                    {
                                        num_vale_vista = NumDoc;
                                        break;
                                    }
                                case "A":
                                    {
                                        num_venta = NumDoc;
                                        break;
                                    }
                            }
                            num_cuotasaux = 1;
                            if (j != num_cuotas)
                            {
                                detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, Convert.ToInt64(Math.Round(montoaux, 0)), moneda, banco, emisor
                                 , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                 , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                 , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                                cheques.Add(detcheq);
                            }
                            else
                            {
                                detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, Convert.ToInt64(Math.Round(montoaux, 0) + montores), moneda, banco, emisor
                                                             , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                                             , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                                             , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                                cheques.Add(detcheq);
                            }
                        }
                    }
                }

                if (DGCheque.Items.Count > 0)
                {
                    DGCheque.ItemsSource = null;
                    DGCheque.Items.Clear();
                }
                else
                {
                    DGCheque.Items.Clear();
                    DGCheque.ItemsSource = null;
                }
                DGCheque.ItemsSource = cheques;

                double MntTotalChq = 0;
                double TotalVPagos = 0;
                var items = new List<MontoMediosdePago>();

                for (int i = items.Count - 1; i >= 0; --i)
                {
                    items.RemoveAt(i);
                }

                if (DGMediosDePagos.Items.Count > 0)
                {
                    DGMediosDePagos.ItemsSource = null;
                    DGMediosDePagos.Items.Clear();
                }
                for (int i = 0; i <= cmbVPMedioPag.Items.Count - 1; i++)
                {
                    try
                    {
                        for (int j = 0; j <= cheques.Count - 1; j++)
                        {
                            if (Convert.ToString(cmbVPMedioPag.Items[i]).Substring(0, 1) == cheques[j].VIA_PAGO)
                                MntTotalChq = MntTotalChq + cheques[j].MONTO;
                        }
                        if (MntTotalChq != 0)
                        {
                            if (moneda == "CLP")
                            {
                                //decimal ValorAux = Convert.ToDecimal(MntTotalChq);
                                string monedachil = Formato.FormatoMoneda(Convert.ToString(MntTotalChq));
                                items.Add(new MontoMediosdePago(Convert.ToString(cmbVPMedioPag.Items[i]), monedachil));
                            }
                            else
                            {
                                //decimal ValorAux = Convert.ToDecimal(MntTotalChq);
                                string monedaforex = Formato.FormatoMonedaExtranjera(Convert.ToString(MntTotalChq));
                                items.Add(new MontoMediosdePago(Convert.ToString(cmbVPMedioPag.Items[i]), monedaforex));
                            }
                            MntTotalChq = 0;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                    }
                }
                // ... Assign ItemsSource of DataGrid.
                for (int i = 0; i <= items.Count - 1; i++)
                {
                    string[] split = items[i].MedioPago.Split(new Char[] { '-' });
                    string MedPago = split[0];

                    if (MedPago == "H " | MedPago == "J " | MedPago == "O ")
                    {
                        switch (MedPago)
                        {
                            case "H ":
                                TotalVPagos = TotalVPagos + Convert.ToDouble(monto);
                                break;
                            case "J ":
                                TotalVPagos = TotalVPagos + Convert.ToDouble(monto);
                                break;
                            case "O ":
                                TotalVPagos = TotalVPagos + Convert.ToDouble(monto);
                                break;
                        }
                    }
                    else
                    {
                        TotalVPagos = TotalVPagos + Convert.ToDouble(items[i].Monto);
                    }
                }
                double MntTotalPend = 0;
                for (int i = 0; i <= partidaseleccionadas.Count - 1; i++)
                {
                    MntTotalPend = MntTotalPend + Convert.ToDouble(partidaseleccionadas[i].MONTOF_PAGAR);
                }
                if (moneda == "CLP")
                {
                    //decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = Formato.FormatoMoneda(Convert.ToString(TotalVPagos));
                    textBlock3.Text = Convert.ToString(monedachil);
                    //decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = Formato.FormatoMoneda(Convert.ToString(MntTotalPend));
                    textBlock4.Text = Convert.ToString(monedachil2);
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = string.Format("{0:0,0}", ValorAux3);

                    if (monedachil3 == "00")
                    {
                        monedachil3 = "0";
                    }

                    if (monedachil3 == "-")
                    {
                        txtVuelto.Text = "";
                    }
                    else
                    {
                        label8.Visibility = Visibility.Visible;
                        textBlock5.Visibility = Visibility.Visible;
                        textBlock5.Text = Convert.ToString(monedachil3);
                        txtVuelto.Text = Convert.ToString(monedachil3);
                        txtVuelto.Visibility = Visibility.Collapsed;
                        lbvuelto.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    //decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = Formato.FormatoMonedaExtranjera(Convert.ToString(TotalVPagos));
                    textBlock3.Text = Convert.ToString(monedachil);
                    //decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = Formato.FormatoMonedaExtranjera(Convert.ToString(MntTotalPend));
                    textBlock4.Text = Convert.ToString(monedachil2);

                    if (MedioPago == "H" | MedioPago == "J" | MedioPago == "O")
                    {
                        decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                        string monedachil3 = Formato.FormatoMonedaExtranjera(Convert.ToString(ValorAux3));
                        if (monedachil3 == "00")
                        {
                            monedachil3 = "0";
                        }
                        textBlock5.Text = Convert.ToString(monedachil3);
                        if (monedachil3.Contains("-"))
                        {
                            txtVuelto.Visibility = Visibility.Visible;
                            lbvuelto.Visibility = Visibility.Visible;
                            txtVuelto.Text = Convert.ToString(monedachil3);
                        }
                        else
                        {
                            txtVuelto.Visibility = Visibility.Collapsed;
                            lbvuelto.Visibility = Visibility.Collapsed;
                            txtVuelto.Text = "";
                        }
                    }
                    else
                    {
                        decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                        string monedachil3 = Formato.FormatoMonedaExtranjera(Convert.ToString(ValorAux3));
                        {
                            monedachil3 = "0";
                        }
                        textBlock5.Text = Convert.ToString(monedachil3);
                    }
                }
                if (chkAbono.IsChecked == true)
                {
                    if (textBlock5.Text != "0")
                    {
                        if (Convert.ToDouble(textBlock3.Text) < Convert.ToDouble(textBlock4.Text))
                        {
                            if (Convert.ToDouble(textBlock5.Text) > 0)
                            {
                                btnConfirPag.IsEnabled = true;
                            }
                        }
                    }
                }
                else
                {
                    if (MedioPago != "H" && MedioPago != "J" && MedioPago != "O")
                    {
                        if (textBlock5.Text == "0")
                        {
                            btnConfirPag.IsEnabled = true;
                        }
                    }
                }
                if (Convert.ToDouble(textBlock5.Text) < 0)
                {
                    if (MedioPago == "H" | MedioPago == "J" | MedioPago == "O")
                    {
                        decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                        txtVuelto.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(ValorAux3));
                        btnValidarEfect.IsEnabled = true;
                        label8.Visibility = Visibility.Collapsed;
                        textBlock5.Visibility = Visibility.Collapsed;

                        DGMediosDePagos.ItemsSource = items;
                        if (items.Count > 0)
                        {
                            DGMediosDePagos.ScrollIntoView(items[items.Count - 1]);
                        }
                    }
                    if (MedioPago != "H" && MedioPago != "J" && MedioPago != "O")
                    {
                        decimal ValorAux3 = Convert.ToDecimal((TotalVPagos) - (MntTotalPend));
                        txtVuelto.IsEnabled = false;
                        txtVuelto.Visibility = Visibility.Collapsed;
                        lbvuelto.Visibility = Visibility.Collapsed;
                        label8.Visibility = Visibility.Collapsed;
                        textBlock5.Visibility = Visibility.Collapsed;
                        if (TotalVPagos > MntTotalPend)
                        {
                            btnConfirPag.IsEnabled = false;
                        }
                        if (TotalVPagos <= MntTotalPend)
                        {
                            btnConfirPag.IsEnabled = true;
                        }
                        DGMediosDePagos.ItemsSource = items;
                        if (items.Count > 0)
                        {
                            DGMediosDePagos.ScrollIntoView(items[items.Count - 1]);
                        }
                    }
                    else
                    {
                        DGMediosDePagos.ItemsSource = items;
                        if (items.Count > 0)
                        {
                            DGMediosDePagos.ScrollIntoView(items[items.Count - 1]);
                        }
                        DPFechVenc.Text = "";
                        cmbBanco.Text = "";
                        txtSucursal.Text = "";
                        txtNumDoc.Text = "";
                        txtMontoFP.Text = "";
                        txtNumCuenta.Text = "";
                        txtRUTGirador.Text = "";
                        txtCantDoc.Text = "";
                        txtNombreGira.Text = "";
                        txtNumVenta.Text = "";
                        cmbIfinan.Text = "";
                        txtObserv.Text = "";
                        txtPatente.Text = "";
                        txtCodOp.Text = "";
                        txtCodAut.Text = "";
                        txtAsig.Text = "";
                        cmbTipoTarjeta.ItemsSource = null;
                        cmbTipoTarjeta.Items.Clear();
                        cmbBanco.ItemsSource = null;
                        cmbBanco.Items.Clear();
                        cmbBancoProp.ItemsSource = null;
                        cmbBancoProp.Items.Clear();
                        cmbCuentasBancosProp.ItemsSource = null;
                        cmbCuentasBancosProp.Items.Clear();
                    }
                }
                else
                {

                    if (MedioPago == "H" | MedioPago == "J" | MedioPago == "O")
                    {
                        //decimal ValorAux3  = Convert.ToDecimal((TotalVPagos) - (MntTotalPend));
                        decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                        txtVuelto.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(ValorAux3));
                        label8.Visibility = Visibility.Visible;
                        textBlock5.Visibility = Visibility.Visible;

                        if (TotalVPagos >= MntTotalPend)
                        {
                            btnConfirPag.IsEnabled = true;
                        }
                        else
                        {
                            btnConfirPag.IsEnabled = false;

                        }
                        DGMediosDePagos.ItemsSource = items;
                        if (items.Count > 0)
                        {
                            DGMediosDePagos.ScrollIntoView(items[items.Count - 1]);
                        }
                    }
                    else
                    {
                        DGMediosDePagos.ItemsSource = items;

                        if (items.Count > 0)
                        {
                            DGMediosDePagos.ScrollIntoView(items[items.Count - 1]);
                        }
                        DPFechVenc.Text = "";
                        cmbBanco.Text = "";
                        txtSucursal.Text = "";
                        txtNumDoc.Text = "";
                        txtMontoFP.Text = "";
                        txtNumCuenta.Text = "";
                        txtRUTGirador.Text = "";
                        txtCantDoc.Text = "";
                        txtNombreGira.Text = "";
                        txtNumVenta.Text = "";
                        cmbIfinan.Text = "";
                        txtObserv.Text = "";
                        txtPatente.Text = "";
                        txtCodOp.Text = "";
                        txtCodAut.Text = "";
                        txtAsig.Text = "";
                        cmbTipoTarjeta.ItemsSource = null;
                        cmbTipoTarjeta.Items.Clear();
                        cmbBanco.ItemsSource = null;
                        cmbBanco.Items.Clear();
                        cmbBancoProp.ItemsSource = null;
                        cmbBancoProp.Items.Clear();
                        cmbCuentasBancosProp.ItemsSource = null;
                        cmbCuentasBancosProp.Items.Clear();
                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }

        //FUNCION QUE TOMA EL INGRESO DE DETALLES DE MEDIOS Y FORMAS DE PAGO, MONTOS Y LLENA EL GRID RESUMEN DE PAGOS (DGCHEQUESMASIVOS).
        private void IngresoFormasDePagoYMontosMasivos(string MedioPago)
        {
            try
            {
                String FechaVenct;
                if (DPFechActualMasivo.Text == "")
                {
                    DPFechActualMasivo.Text = Calendario.Text;
                    DPFechActualMasivo.Text = Convert.ToString(Calendario.SelectedDate.Value);
                }

                if (DPFechVencMasivo.Text != "")
                {
                    FechaVenct = DPFechVencMasivo.Text.Substring(0, 10);
                }
                else
                {
                    FechaVenct = DPFechActualMasivo.Text.Substring(0, 10);
                }
                string mandt = "";
                string id_caja = Convert.ToString(textBlock6.Content);
                string land = Convert.ToString(lblPais.Content);
                string id_comprobante = "";
                string id_detalle = "";
                string via_pago = MedioPago;
                string Valor = txtMontoFPMasivo.Text.Replace(".", "");
                Valor = Valor.Replace(",", ".");
                double monto = Convert.ToDouble(Valor);
                string moneda = cmbMonedaMasivo.Text;
                string[] split = cmbBancoMasivo.Text.Split(new Char[] { '-' });
                string banco = split[0];
                string emisor = txtRUTGiradorMasivo.Text;
                string num_cheque = "";
                string cod_autorizacion = txtCodAutMasivo.Text;
                int num_cuotas;
                int num_cuotasaux;
                if (txtCantDocMasivo.Text == "")
                {
                    txtCantDocMasivo.Text = "0";
                }

                num_cuotas = Convert.ToInt16(txtCantDocMasivo.Text);

                string fecha_venc = FechaVenct;
                string texto_posicion = txtObservMasivo.Text;
                string anexo = "";
                string sucursal = txtSucursalMasivo.Text;
                string num_cuenta = txtNumCuentaMasivo.Text;
                string num_tarjeta = "";
                string num_vale_vista = "";
                string patente = txtPatenteMasivo.Text;
                string num_venta = txtNumVentaMasivo.Text;
                string pagare = "";
                string fecha_emision = DPFechActualMasivo.Text;
                string nombre_girador = txtNombreGiraMasivo.Text;
                string carta_curse = "";
                string num_transfer = "";
                string num_deposito = "";
                string cta_banco = cmbCuentasBancosPropMasivo.Text;
                string ifinan = "";
                string corre = "";
                if (cmbIfinanMasivo.Text != "")
                {
                    int posicion = cmbIfinanMasivo.Text.LastIndexOf("-");
                    posicion = posicion - 1;
                    ifinan = cmbIfinanMasivo.Text.Substring(0, posicion);
                }
                else
                {
                    ifinan = "";
                }
                string zuonr = txtAsigMasivo.Text;
                string hkont = "";
                string prctr = "";
                string znop = txtCodOpMasivo.Text;
                string NumDoc = txtNumDocMasivo.Text;
                string NumCtaCte = txtNumCuentaMasivo.Text;
                string Patente = txtPatenteMasivo.Text;
                if (cmbIntervaloMasivo.Text == "")
                {
                    cmbIntervaloMasivo.Text = "0";
                }
                int Intervalo = Convert.ToInt16(cmbIntervaloMasivo.Text);

                VIAS_PAGO_MASIVO detcheqMasivo;
                if (num_cuotas < 2) //VIAS DE PAGO SIN CUOTAS
                {
                    //DEPENDIENDO DEL MEDIO DE PAGO SE HACE LA LOGICA PARA INCORPORAR EL NUMERO DEL MEDIO DE PAGO AL CAMPO CORRECTO DE LA RFC
                    switch (MedioPago)
                    {
                        case "B":
                            {
                                num_deposito = txtNumDocMasivo.Text;
                                break;
                            }
                        case "D":
                            {
                                num_deposito = txtNumDocMasivo.Text;
                                break;
                            }
                        case "F":
                            {
                                num_cheque = txtNumDocMasivo.Text;
                                break;
                            }
                        case "G":
                            {
                                num_cheque = txtNumDocMasivo.Text;
                                break;
                            }
                        case "K":
                            {
                                carta_curse = txtNumDocMasivo.Text;
                                num_venta = NumDoc;
                                break;
                            }
                        case "M":
                            {
                                pagare = txtNumDocMasivo.Text;
                                break;
                            }
                        case "P":
                            {
                                pagare = txtNumDocMasivo.Text;
                                break;
                            }
                        case "R":
                            {
                                num_tarjeta = txtNumDocMasivo.Text;
                                break;
                            }
                        case "S":
                            {
                                num_tarjeta = txtNumDocMasivo.Text;
                                break;
                            }
                        case "U":
                            {
                                num_transfer = txtNumDocMasivo.Text;
                                break;
                            }
                        case "V":
                            {
                                num_vale_vista = txtNumDocMasivo.Text;
                                break;
                            }
                        case "A":
                            {
                                num_venta = NumDoc;
                                break;
                            }
                    }
                    detcheqMasivo = new VIAS_PAGO_MASIVO(mandt, land, id_comprobante, id_detalle, id_caja, via_pago, monto, moneda, banco, emisor
                        , num_cheque, cod_autorizacion, Convert.ToString(num_cuotas), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                        , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                        , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                    chequesMasiv.Add(detcheqMasivo);
                }
                else
                {
                    double montotot = 0;
                    double montores = 0;
                    int j = 0;
                    if (monto % num_cuotas == 0) //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS EXACTO
                    {
                        for (int i = 1; i <= num_cuotas; i++)
                        {
                            double montoaux = Convert.ToDouble(monto / num_cuotas);
                            j++;

                            if (i != 1) //CALCULO DE LAS FECHAS DE VENCIMIENTO DE ACUERDO AL NUMERO DE CUOTAS
                            {
                                DateTime FechaVenctAux = Convert.ToDateTime(FechaVenct);
                                FechaVenctAux = FechaVenctAux.AddDays(Intervalo);
                                FechaVenct = Convert.ToString(FechaVenctAux);
                                fecha_venc = Convert.ToString(FechaVenctAux.ToString("dd/MM/yyyy"));
                                if (MedioPago != "S")
                                {
                                    NumDoc = Convert.ToString(Convert.ToInt64(NumDoc) + 1);
                                }
                            }
                            //DEPENDIENDO DEL MEDIO DE PAGO SE HACE LA LOGICA PARA INCORPORAR EL NUMERO DEL MEDIO DE PAGO AL CAMPO CORRECTO DE LA RFC
                            switch (MedioPago)
                            {
                                case "B":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "D":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "F":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "G":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "K":
                                    {
                                        num_venta = NumDoc;
                                        carta_curse = NumDoc;
                                        break;
                                    }
                                case "M":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "P":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "R":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "S":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "U":
                                    {
                                        num_transfer = NumDoc;
                                        break;
                                    }
                                case "V":
                                    {
                                        num_vale_vista = NumDoc;
                                        break;
                                    }
                                case "A":
                                    {
                                        num_venta = NumDoc;
                                        break;
                                    }
                            }
                            num_cuotasaux = 1;

                            detcheqMasivo = new VIAS_PAGO_MASIVO(mandt, land, id_comprobante, id_detalle, id_caja, via_pago, Convert.ToInt64(Math.Round(montoaux, 0) + montores), moneda, banco, emisor
                                                           , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                                           , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                                           , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                            chequesMasiv.Add(detcheqMasivo);
                        }
                    }
                    else //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS CON UN RESIDUO QUE SE SUMA EN LA CUOTA FINAL
                    {
                        for (int i = 1; i <= num_cuotas; i++)
                        {
                            double montoaux = Convert.ToDouble(monto / num_cuotas);
                            montotot = montotot + Math.Round(montoaux, 0);
                            montores = (monto) - montotot;
                            j++;
                            if (i != 1) //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS EXACTO
                            {
                                DateTime FechaVenctAux = Convert.ToDateTime(FechaVenct);
                                FechaVenctAux = FechaVenctAux.AddDays(Intervalo);
                                FechaVenct = Convert.ToString(FechaVenctAux);
                                fecha_venc = Convert.ToString(FechaVenctAux);
                                if (MedioPago != "S")
                                {
                                    NumDoc = Convert.ToString(Convert.ToInt64(NumDoc) + 1);
                                }
                            }
                            //DEPENDIENDO DEL MEDIO DE PAGO SE HACE LA LOGICA PARA INCORPORAR EL NUMERO DEL MEDIO DE PAGO AL CAMPO CORRECTO DE LA RFC
                            switch (MedioPago)
                            {
                                case "B":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "D":
                                    {
                                        num_deposito = NumDoc;
                                        break;
                                    }
                                case "F":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "G":
                                    {
                                        num_cheque = NumDoc;
                                        break;
                                    }
                                case "K":
                                    {
                                        carta_curse = NumDoc;
                                        num_venta = NumDoc;
                                        break;
                                    }
                                case "M":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "P":
                                    {
                                        pagare = NumDoc;
                                        break;
                                    }
                                case "R":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "S":
                                    {
                                        num_tarjeta = NumDoc;
                                        break;
                                    }
                                case "U":
                                    {
                                        num_transfer = NumDoc;
                                        break;
                                    }
                                case "V":
                                    {
                                        num_vale_vista = NumDoc;
                                        break;
                                    }
                                case "A":
                                    {
                                        num_venta = NumDoc;
                                        break;
                                    }
                            }
                            num_cuotasaux = 1;

                            if (j != num_cuotas)
                            {
                                detcheqMasivo = new VIAS_PAGO_MASIVO(mandt, land, id_comprobante, id_detalle, via_pago, id_caja, Convert.ToInt64(Math.Round(montoaux, 0)), moneda, banco, emisor
                                 , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                 , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                 , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                                chequesMasiv.Add(detcheqMasivo);
                            }
                            else
                            {
                                detcheqMasivo = new VIAS_PAGO_MASIVO(mandt, land, id_comprobante, id_detalle, id_caja, via_pago, Convert.ToInt64(Math.Round(montoaux, 0) + montores), moneda, banco, emisor
                                                             , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                                             , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                                             , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                                chequesMasiv.Add(detcheqMasivo);
                            }
                        }
                    }
                }
                if (DGChequeMasivo.Items.Count > 0)
                {
                    DGChequeMasivo.ItemsSource = null;
                    DGChequeMasivo.Items.Clear();
                }
                else
                {
                    DGChequeMasivo.Items.Clear();
                    DGChequeMasivo.ItemsSource = null;
                }

                DGChequeMasivo.ItemsSource = chequesMasiv;

                double MntTotalChq = 0;
                double TotalVPagos = 0;
                var items = new List<MontoMediosdePago>();

                for (int i = items.Count - 1; i >= 0; --i)
                {
                    items.RemoveAt(i);
                }

                if (DGMediosDePagosMasivo.Items.Count > 0)
                {
                    DGMediosDePagosMasivo.ItemsSource = null;
                    DGMediosDePagosMasivo.Items.Clear();
                }

                for (int i = 0; i <= cmbVPMedioPagMasivo.Items.Count - 1; i++)
                {
                    try
                    {
                        for (int j = 0; j <= chequesMasiv.Count - 1; j++)
                        {
                            if (Convert.ToString(cmbVPMedioPagMasivo.Items[i]).Substring(0, 1) == chequesMasiv[j].VIA_PAGO)
                                MntTotalChq = MntTotalChq + Convert.ToDouble(chequesMasiv[j].MONTO);
                        }
                        if (MntTotalChq != 0)
                        {
                            if (moneda == "CLP")
                            {
                                //decimal ValorAux = Convert.ToDecimal(MntTotalChq);
                                string monedachil = Formato.FormatoMoneda(Convert.ToString(MntTotalChq));
                                items.Add(new MontoMediosdePago(Convert.ToString(cmbVPMedioPagMasivo.Items[i]), monedachil));
                            }
                            else
                            {
                                //decimal ValorAux = Convert.ToDecimal(MntTotalChq);
                                string monedaforex = Formato.FormatoMonedaExtranjera(Convert.ToString(MntTotalChq));
                                items.Add(new MontoMediosdePago(Convert.ToString(cmbVPMedioPagMasivo.Items[i]), monedaforex));
                            }
                            MntTotalChq = 0;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                    }
                }
                for (int i = 0; i <= items.Count - 1; i++)
                {
                    TotalVPagos = TotalVPagos + Convert.ToDouble(items[i].Monto);
                }
                double MntTotalPend = 0;

                MntTotalPend = MntTotalPend + Convert.ToDouble(textBlock4Masivo.Text);

                if (moneda == "CLP")
                {
                    //decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = Formato.FormatoMoneda(Convert.ToString(TotalVPagos));
                    textBlock3Masivo.Text = Convert.ToString(monedachil);
                    //decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = Formato.FormatoMoneda(Convert.ToString(MntTotalPend));
                    textBlock4Masivo.Text = Convert.ToString(monedachil2);
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = Formato.FormatoMoneda(Convert.ToString(ValorAux3));
                    if (monedachil3 == "00")
                    {
                        monedachil3 = "0";
                    }
                    textBlock5Masivo.Text = Convert.ToString(monedachil3);
                }
                else
                {
                    //decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = Formato.FormatoMonedaExtranjera(Convert.ToString(TotalVPagos));
                    textBlock3Masivo.Text = Convert.ToString(monedachil);
                    //decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = Formato.FormatoMonedaExtranjera(Convert.ToString(MntTotalPend));
                    textBlock4Masivo.Text = Convert.ToString(monedachil2);
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = Formato.FormatoMonedaExtranjera(Convert.ToString(ValorAux3));
                    textBlock5Masivo.Text = Convert.ToString(monedachil3);
                    if (monedachil3 == "00")
                    {
                        monedachil3 = "0";
                    }
                    textBlock5Masivo.Text = Convert.ToString(monedachil3);
                }

                if (chkAbono.IsChecked == true)
                {
                    if (textBlock5Masivo.Text != "0")
                    {
                        if (Convert.ToDouble(textBlock3Masivo.Text) < Convert.ToDouble(textBlock4Masivo.Text))
                        {
                            if (Convert.ToDouble(textBlock5Masivo.Text) > 0)
                            {
                                btnConfirPagMasivo.IsEnabled = true;
                            }
                        }
                    }
                }
                else
                {
                    if (textBlock5Masivo.Text == "0")
                    {
                        btnConfirPagMasivo.IsEnabled = true;
                    }
                }
                if (Convert.ToDouble(textBlock5Masivo.Text) < 0)
                {
                    System.Windows.Forms.MessageBox.Show("Montos de vias de pago es superior a la cantidad a cancelar");
                    txtMontoFPMasivo.Text = "";
                    textBlock3Masivo.Text = "";
                    textBlock5Masivo.Text = "";
                    items.Clear();
                    cheques.Clear();
                }
                else
                {
                    DGMediosDePagosMasivo.ItemsSource = items;
                    if (items.Count > 0)
                    {
                        DGMediosDePagosMasivo.ScrollIntoView(items[items.Count - 1]);
                    }
                    DPFechVencMasivo.Text = "";
                    cmbBancoMasivo.Text = "";
                    txtSucursalMasivo.Text = "";
                    txtNumDocMasivo.Text = "";
                    txtMontoFPMasivo.Text = "";
                    txtNumCuentaMasivo.Text = "";
                    txtRUTGiradorMasivo.Text = "";
                    txtCantDocMasivo.Text = "";
                    txtNombreGiraMasivo.Text = "";
                    txtNumVentaMasivo.Text = "";
                    cmbIfinanMasivo.Text = "";
                    txtObservMasivo.Text = "";
                    txtPatenteMasivo.Text = "";
                    txtCodOpMasivo.Text = "";
                    txtCodAutMasivo.Text = "";
                    txtAsigMasivo.Text = "";
                    cmbTipoTarjetaMasivo.ItemsSource = null;
                    cmbTipoTarjetaMasivo.Items.Clear();
                    cmbBancoMasivo.ItemsSource = null;
                    cmbBancoMasivo.Items.Clear();
                    cmbBancoPropMasivoMasivo.ItemsSource = null;
                    cmbBancoPropMasivoMasivo.Items.Clear();
                    cmbCuentasBancosPropMasivo.ItemsSource = null;
                    cmbCuentasBancosPropMasivo.Items.Clear();
                    if (textBlock5Masivo.Text == "0")
                    {
                        btnBuscarPM.Visibility = Visibility.Visible;
                        PrgBarExcel.Visibility = Visibility.Visible;
                        btnAgregaMtoMasivo.Visibility = Visibility.Collapsed;
                        textBlock5Masivo.IsEnabled = false;
                        textBlock4Masivo.IsEnabled = false;
                        textBlock3Masivo.IsEnabled = false;
                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }

        //FUNCION QUE TRAE EL DETALLE DE LOS REGISTROS DE LA GRILLA DE LAS PARTIDAS ABIERTAS Y/O NOTAS DE VENTAS Y LO CARGA EN UNA NUEVA VENTANA
        private void DGPagos_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                try
                {
                    partidaseleccionadas = new List<T_DOCUMENTOS>();
                    partidaseleccionadas.Clear();
                    partidaseleccionadas.Add(DGPagos.SelectedItem as T_DOCUMENTOS);

                    DetalleDocumentos detalle = new DetalleDocumentos(partidaseleccionadas[0].NDOCTO, partidaseleccionadas[0].NREF
                        , partidaseleccionadas[0].RUTCLI, partidaseleccionadas[0].COD_CLIENTE, partidaseleccionadas[0].NOMCLI
                        , partidaseleccionadas[0].CEBE, partidaseleccionadas[0].CONTROL_CREDITO, partidaseleccionadas[0].SOCIEDAD
                        , partidaseleccionadas[0].FECHA_DOC, partidaseleccionadas[0].FECVENCI, partidaseleccionadas[0].DIAS_ATRASO
                        , partidaseleccionadas[0].MONEDA, partidaseleccionadas[0].CLASE_DOC, partidaseleccionadas[0].CLASE_CUENTA
                        , partidaseleccionadas[0].CME, partidaseleccionadas[0].ACC, partidaseleccionadas[0].ESTADO
                        , partidaseleccionadas[0].COND_PAGO, partidaseleccionadas[0].MONTOF_PAGAR, partidaseleccionadas[0].MONTOF_ABON
                        , partidaseleccionadas[0].MONTOF);
                    detalle.Show();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                }
                this.Topmost = false;
            }


            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
            }
            GC.Collect();
        }
        //FUNCION QUE TRAE LOS DOCUMENTOS DESDE EL MONITOR
        private void ListaDocumentosPendientesDesdeMonitor()
        {
            try
            {
                timer.Stop();
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }
                DGPagos.ItemsSource = null;
                DGPagos.Items.Clear();

                if (DatPckPgDoc.Text == "")
                {
                    DatPckPgDoc.Text = Calendario.Text;
                }
                try
                {
                    monitorseleccionado.Clear();
                    if (this.DGMonitor.SelectedItems.Count > 0)
                        for (int i = 0; i < DGMonitor.SelectedItems.Count; i++)
                        {
                            {
                                monitorseleccionado.Add(DGMonitor.SelectedItems[i] as T_DOCUMENTOS);
                            }

                        }

                    DGPagos.ItemsSource = null;
                    DGPagos.Items.Clear();
                    DGPagos.ItemsSource = monitorseleccionado;
                    DGPagos.SelectAll();
                    partidaseleccionadas = new List<T_DOCUMENTOS>();

                    int posicion = 0;
                    if (this.DGPagos.SelectedItems.Count > 0)
                    {
                        for (int i = 0; i < DGPagos.SelectedItems.Count; i++)
                        {
                            {
                                partidaseleccionadas.Add(DGPagos.SelectedItems[i] as T_DOCUMENTOS);
                            }
                        }
                        double Monto = 0;

                        for (int i = 0; i < monitorseleccionado.Count; i++)
                        {
                            if (monitorseleccionado[i].MONTOF == "")
                            {
                                monitorseleccionado[i].MONTOF = "0";
                            }
                            monitorseleccionado[i].MONTOF = monitorseleccionado[i].MONTOF.Trim();
                            if (monitorseleccionado[i].MONTOF.Contains("-"))
                            {
                                posicion = monitorseleccionado[i].MONTOF.IndexOf("-");
                                if (posicion == monitorseleccionado[i].MONTOF.Length - 1)
                                {
                                    monitorseleccionado[i].MONTOF = monitorseleccionado[i].MONTOF.Substring(posicion, 1) + monitorseleccionado[i].MONTOF.Substring(0, posicion);
                                }
                            }
                            Monto = Monto + Convert.ToDouble(monitorseleccionado[i].MONTOF);
                        }
                        textBlock4.Text = Convert.ToString(Monto);
                        GBViasPago.HorizontalAlignment = System.Windows.HorizontalAlignment.Stretch;
                        GBViasPago.Margin = new Thickness(1, 406, 6, 0);
                        GBViasPago.VerticalAlignment = VerticalAlignment.Top;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);

                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        //FUNCION QUE TRAE LOS DOCUMENTOS PENDIENTES O PARTIDAS ABIERTAS POR CARGAS MASIVAS 
        private void ListaDocumentosPendientesCargasMasivas(string thisFileName)
        {
            try
            {
                string RutExc = "";
                string SocExc = "";
                List<PagosMasivosNuevo> ListaExc = new List<PagosMasivosNuevo>();
                List<VIAS_PAGO_MASIVO> ListViasPagos = new List<VIAS_PAGO_MASIVO>();
                VIAS_PAGO_MASIVO p_viasPagoMasivos;

                //*RFC PAGO DE DOCUMENTOS
                // List<DetalleViasPago> ListViasPagos = new List<DetalleViasPago>();

                for (int i = 1; i <= DGChequeMasivo.Items.Count; i++)
                {
                    if (i == 1)
                    {
                        DGChequeMasivo.Items.MoveCurrentToFirst();
                    }
                    ListViasPagos.Add(DGChequeMasivo.Items.CurrentItem as VIAS_PAGO_MASIVO);

                    DGChequeMasivo.Items.MoveCurrentToNext();
                }
                //LLAMADO A LA FUNCION QUE LEE EL ARCHIVO EXCEL
                RecogerDatosExcel2(thisFileName, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ref SocExc, ref RutExc, out ListaExc, ref  PrgBarExcel);

                //RFC que hace la compensacion de pagos masivos
                PagosMasivosNew pagosmasivos = new PagosMasivosNew();
                pagosmasivos.pagosmasivos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                    , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content)
                    , Convert.ToString(DateTime.Today), thisFileName, logApertura2[0].ID_REGISTRO, logApertura2[0].ID_CAJA, logApertura2[0].MONEDA, ListaExc, ListViasPagos);

                if (pagosmasivos.message != "")
                {
                    System.Windows.Forms.MessageBox.Show(pagosmasivos.message);

                    if (pagosmasivos.comprobante != "")
                    {
                        ImpresionesDeDocumentosAutomaticas(pagosmasivos.comprobante, "X");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No se generó comprobante de pago");
                    }

                    LimpiarPagMasivo();
                }
                if (pagosmasivos.errormessage != "")
                {
                    System.Windows.Forms.MessageBox.Show(pagosmasivos.errormessage);
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
            GBViasPagoMasivos.Visibility = Visibility.Visible;
        }

        //FUNCION QUE TRAE LOS DOCUMENTOS A PAGAR A PARTIR DE UN ARCHIVO EXCEL
        static void RecogerDatosExcel2(string ruta, string usuario, string sucursal, string idcaja, ref string SocExc, ref string RutExc, out List<PagosMasivosNuevo> ListaExc, ref System.Windows.Controls.ProgressBar PrgBarExcel)
        {
            ListaExc = new List<PagosMasivosNuevo>();
            Microsoft.Office.Interop.Excel._Application xlApp;
            Microsoft.Office.Interop.Excel._Workbook xlLibro;
            Microsoft.Office.Interop.Excel._Worksheet xlHoja1;
            Microsoft.Office.Interop.Excel.Sheets xlHojas;
            string fileName = ruta;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlLibro = xlApp.Workbooks.Open(fileName);
            try
            {
                xlHojas = xlLibro.Sheets;
                try
                {
                    int k = 1;
                    //Asignamos la hoja con la que queremos trabajar: 
                    xlHoja1 = (Microsoft.Office.Interop.Excel._Worksheet)xlHojas["PagoMasivoCliente"];
                    int n = xlHoja1.UsedRange.Rows.Count;
                    PrgBarExcel.Maximum = n;
                    int j = 4;
                    int m = 2;
                    int l = 1;
                    int verificador = 0;
                    SocExc = "";
                    RutExc = "";

                    PrgBarExcel.Value = 0;
                    string row = "";
                    string col = "";
                    string value = "";

                    PagosMasivosNuevo pagosm = new PagosMasivosNuevo(row, col, value);
                    col = "2";
                    while (col.Length < 4)
                    {
                        col = "0" + col;
                    }
                    row = "1";
                    while (row.Length < 4)
                    {
                        row = "0" + row;
                    }
                    value = (string)xlHoja1.Cells["1", "B"].Text;
                    pagosm.COL = col;
                    pagosm.ROW = row;
                    pagosm.VALUE = value;
                    ListaExc.Add(pagosm);
                    pagosm = new PagosMasivosNuevo(row, col, value);
                    col = "4";
                    while (col.Length < 4)
                    {
                        col = "0" + col;
                    }
                    row = "1";
                    while (row.Length < 4)
                    {
                        row = "0" + row;
                    }
                    value = (string)xlHoja1.Cells["1", "D"].Text;
                    pagosm.COL = col;
                    pagosm.ROW = row;
                    pagosm.VALUE = value;
                    ListaExc.Add(pagosm);
                    pagosm = new PagosMasivosNuevo(row, col, value);
                    col = "2";
                    while (col.Length < 4)
                    {
                        col = "0" + col;
                    }
                    row = "2";
                    while (row.Length < 4)
                    {
                        row = "0" + row;
                    }
                    value = (string)xlHoja1.Cells["2", "B"].Text;
                    pagosm.COL = col;
                    pagosm.ROW = row;
                    pagosm.VALUE = value;
                    ListaExc.Add(pagosm);
                    for (int i = 3; i <= n; i++)
                    {
                        if (verificador >= 2)
                        {
                            break;
                        }

                        if (((string)xlHoja1.Cells[j, "A"].Text != "") && ((string)xlHoja1.Cells[j, "B"].Text != ""))
                        {
                            for (int r = 1; r <= 2; r++)
                            {
                                row = Convert.ToString(j);
                                while (row.Length < 4)
                                {
                                    row = "0" + row;
                                }
                                col = Convert.ToString(r);
                                while (col.Length < 4)
                                {
                                    col = "0" + col;
                                }
                                if (r == 1)
                                    value = (string)xlHoja1.Cells[j, "A"].Text;
                                else
                                    value = (string)xlHoja1.Cells[j, "B"].Text;
                                pagosm = new PagosMasivosNuevo(row, col, value);
                                pagosm.COL = col;
                                pagosm.ROW = row;
                                pagosm.VALUE = value;
                                if (value != "")
                                {
                                    ListaExc.Add(pagosm);
                                }
                            }
                            j++;
                            PrgBarExcel.Value = (PrgBarExcel.Value + 1);
                        }
                        else
                        {
                            verificador++;
                        }
                    }
                    SocExc = (string)xlHoja1.Cells[2, "B"].Text;
                    RutExc = (string)xlHoja1.Cells[1, "B"].Text;

                    pagosm = new PagosMasivosNuevo(row, col, value);
                    col = "2";
                    while (col.Length < 4)
                    {
                        col = "0" + col;
                    }
                    row = Convert.ToString(xlHoja1.UsedRange.Rows.Count);
                    while (row.Length < 4)
                    {
                        row = "0" + row;
                    }

                    value = (string)xlHoja1.Cells[Convert.ToString(xlHoja1.UsedRange.Rows.Count), "B"].Text;
                    pagosm.COL = col;
                    pagosm.ROW = row;
                    pagosm.VALUE = value;
                    ListaExc.Add(pagosm);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    logCaja.EscribeLogCaja(System.DateTime.Now, usuario, idcaja, sucursal, ex.Message + ex.StackTrace);
                }
                if (ListaExc.Count == 0)
                {
                    System.Windows.MessageBox.Show("Error en el archivo excel a cargar. Revise formato del archivo o el formato de la plantilla o si el archivo tiene datos");
                }
                else
                {
                    System.Windows.MessageBox.Show(Convert.ToString(ListaExc.Count) + " documentos cargados");
                }
            }

            finally
            {
                xlLibro.Close(false);
                xlApp.Quit();
                PrgBarExcel.Value = 0;
                GC.Collect();
            }
        }

        static void RecogerDatosExcel(string ruta, string usuario, string sucursal, string idcaja, ref string SocExc, ref string RutExc, out List<PagosMasivos> ListaExc, ref System.Windows.Controls.ProgressBar PrgBarExcel)
        {
            ListaExc = new List<PagosMasivos>();

            Microsoft.Office.Interop.Excel._Application xlApp;
            Microsoft.Office.Interop.Excel._Workbook xlLibro;
            Microsoft.Office.Interop.Excel._Worksheet xlHoja1;
            Microsoft.Office.Interop.Excel.Sheets xlHojas;
            string fileName = ruta;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            xlLibro = xlApp.Workbooks.Open(fileName);
            try
            {
                xlHojas = xlLibro.Sheets;
                try
                {
                    int k = 1;
                    xlHoja1 = (Microsoft.Office.Interop.Excel._Worksheet)xlHojas["Hoja1"];
                    int n = xlHoja1.UsedRange.Rows.Count;
                    PrgBarExcel.Maximum = n;
                    int j = 4;
                    int m = 2;
                    int l = 1;
                    int verificador = 0;
                    SocExc = "";
                    RutExc = "";
                    PrgBarExcel.Value = 0;
                    for (int i = 3; i <= n; i++)
                    {
                        if (verificador >= 2)
                        {

                            break;
                        }
                        if (((string)xlHoja1.Cells[j, "A"].Text != "") && ((string)xlHoja1.Cells[j, "B"].Text != "") && ((string)xlHoja1.Cells[j, "C"].Text != ""))
                        {
                            string referencia = (string)xlHoja1.Cells[j, "A"].Text;
                            string monto = (string)xlHoja1.Cells[j, "B"].Text;
                            string moneda = (string)xlHoja1.Cells[j, "C"].Text;
                            PagosMasivos pagosm = new PagosMasivos(referencia, monto, moneda);
                            ListaExc.Add(pagosm);
                            SocExc = (string)xlHoja1.Cells[m, "B"].Text;
                            RutExc = (string)xlHoja1.Cells[l, "B"].Text;
                            j++;
                            PrgBarExcel.Value = (PrgBarExcel.Value + 1);
                        }
                        else
                        {
                            verificador++;
                        }
                    }
                    SocExc = (string)xlHoja1.Cells[m, "B"].Text;
                    RutExc = (string)xlHoja1.Cells[l, "B"].Text;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    logCaja.EscribeLogCaja(System.DateTime.Now, usuario, idcaja, sucursal, ex.Message + ex.StackTrace);
                }
                if (ListaExc.Count == 0)
                {
                    System.Windows.MessageBox.Show("Error en el archivo excel a cargar. Revise formato del archivo o el formato de la plantilla o si el archivo tiene datos");
                }
                else
                {
                    System.Windows.MessageBox.Show(Convert.ToString(ListaExc.Count) + " documentos cargados");
                }
            }
            finally
            {
                xlLibro.Close(false);
                xlApp.Quit();
                PrgBarExcel.Value = 0;
                GC.Collect();
            }
        }
        //FUNCION QUE TRAE LOS DOCUMENTOS A PAGAR A PARTIR DE LA BUSQUEDA POR RUT O DOCUMENTO
        private void ListaDocumentosPendientes()
        {
            try
            {
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }

                partidasabiertas.ObjDatosPartidasOpen.Clear();
                DGPagos.ItemsSource = null;
                DGPagos.Items.Clear();
                if (DatPckPgDoc.Text == "")
                {
                    DatPckPgDoc.Text = Calendario.Text;
                }
                if (RBRut.IsChecked == true)
                {
                    string RUTAux = "";
                    foreach (char value in txtRut.Text.ToUpper())
                    {
                        if (value != 8207)
                        {
                            RUTAux = RUTAux + value;
                        }
                    }

                    RUTAux = RUTAux.Trim();
                    //***RFC Partidas abiertas para pago busqueda por RUT
                    String RUT = DigitoVerificador(RUTAux);
                    //Verificacion de RUT
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRut.Focus();
                    }
                    else
                    {
                        partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "", txtDocu.Text.ToUpper(), txtRut.Text.ToUpper(), Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text), Convert.ToString(lblPais.Content), "", "RUT");
                    }
                }
                else if (RBDoc.IsChecked == true)
                {
                    string Documento = "";
                    Documento = txtDocu.Text.ToUpper();
                    if (Documento.Contains("-"))
                    {
                        System.Windows.Forms.MessageBox.Show("Introduzca un número válido de documento/comprobante");
                    }
                    else
                    {
                        while (Documento.Length < 10)
                        {
                            Documento = "0" + Documento;
                        }
                        txtDocu.Text = Documento.ToUpper();

                        partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "", txtDocu.Text, txtRut.Text, Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text), Convert.ToString(lblPais.Content), "", "Documento");
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }

                if (partidasabiertas.ObjDatosPartidasOpen.Count > 0)
             {
                    int COUNT = 1;
                    GBDocsAPagar.Visibility = Visibility.Visible;
                    DGPagos.ItemsSource = null;
                    DGPagos.Items.Clear();
                    List<T_DOCUMENTOSAUX> partidaopen = new List<T_DOCUMENTOSAUX>();

                    for (int k = 0; k < partidasabiertas.ObjDatosPartidasOpen.Count; k++)
                    {
                        T_DOCUMENTOSAUX partOpen = new T_DOCUMENTOSAUX();

                        partOpen.ID = Convert.ToString(COUNT);
                        partOpen.ISSELECTED = false;
                        partOpen.ACC = partidasabiertas.ObjDatosPartidasOpen[k].ACC;
                        partOpen.CEBE = partidasabiertas.ObjDatosPartidasOpen[k].CEBE;
                        partOpen.CLASE_CUENTA = partidasabiertas.ObjDatosPartidasOpen[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = partidasabiertas.ObjDatosPartidasOpen[k].CLASE_DOC;
                        partOpen.CME = partidasabiertas.ObjDatosPartidasOpen[k].CME;
                        partOpen.COD_CLIENTE = partidasabiertas.ObjDatosPartidasOpen[k].COD_CLIENTE;
                        partOpen.COND_PAGO = partidasabiertas.ObjDatosPartidasOpen[k].COND_PAGO;
                        partOpen.CONTROL_CREDITO = partidasabiertas.ObjDatosPartidasOpen[k].CONTROL_CREDITO;
                        partOpen.DIAS_ATRASO = partidasabiertas.ObjDatosPartidasOpen[k].DIAS_ATRASO;
                        partOpen.ESTADO = partidasabiertas.ObjDatosPartidasOpen[k].ESTADO;
                        partOpen.FECHA_DOC = partidasabiertas.ObjDatosPartidasOpen[k].FECHA_DOC;
                        partOpen.FECVENCI = partidasabiertas.ObjDatosPartidasOpen[k].FECVENCI;
                        partOpen.ICONO = partidasabiertas.ObjDatosPartidasOpen[k].ICONO;
                        partOpen.MONEDA = partidasabiertas.ObjDatosPartidasOpen[k].MONEDA;
                        partOpen.MONTO = partidasabiertas.ObjDatosPartidasOpen[k].MONTO;
                        partOpen.MONTOF = partidasabiertas.ObjDatosPartidasOpen[k].MONTOF;
                        partOpen.MONTO_ABONADO = partidasabiertas.ObjDatosPartidasOpen[k].MONTO_ABONADO;
                        partOpen.MONTOF_ABON = partidasabiertas.ObjDatosPartidasOpen[k].MONTOF_ABON;
                        partOpen.MONTO_PAGAR = partidasabiertas.ObjDatosPartidasOpen[k].MONTO_PAGAR;
                        partOpen.MONTOF_PAGAR = partidasabiertas.ObjDatosPartidasOpen[k].MONTOF_PAGAR;
                        partOpen.NDOCTO = partidasabiertas.ObjDatosPartidasOpen[k].NDOCTO;
                        partOpen.NOMCLI = partidasabiertas.ObjDatosPartidasOpen[k].NOMCLI;
                        partOpen.NREF = partidasabiertas.ObjDatosPartidasOpen[k].NREF;
                        partOpen.RUTCLI = partidasabiertas.ObjDatosPartidasOpen[k].RUTCLI;
                        partOpen.SOCIEDAD = partidasabiertas.ObjDatosPartidasOpen[k].SOCIEDAD;
                        partidaopen.Add(partOpen);
                        COUNT++;
                    }
                    DGPagos.ItemsSource = partidaopen;
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        //FUNCION QUE TRAE LOS DOCUMENTOS A PAGAR POR ANTICIPOS A PARTIR DE LA BUSQUEDA POR RUT O DOCUMENTO
        private void ListaDocumentosPendientesAnticipos()
        {
            try
            {
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }

                anticipos.ObjDatosAnticipos.Clear();
                DGPagos.ItemsSource = null;
                DGPagos.Items.Clear();

                if (DatPckPgDoc.Text == "")
                {
                    DatPckPgDoc.Text = Calendario.Text;
                }

                if (RBRUTAnt.IsChecked == true)
                {
                    string RUTAux = "";
                    foreach (char value in txtRUTAnt.Text.ToUpper())
                    {
                        if (value != 8207)
                        {
                            RUTAux = RUTAux + value;
                        }
                    }
                    RUTAux = RUTAux.Trim();
                    //***RFC Anticipos para pago busqueda por RUT
                    String RUT = DigitoVerificador(RUTAux);
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRUTAnt.Focus();
                    }
                    else
                    {
                        anticipos.anticiposopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocuAnt.Text, txtRUTAnt.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "RUT");
                    }
                }
                else if (RBDocuAnt.IsChecked == true)
                {
                    string Documento = "";
                    Documento = txtDocuAnt.Text;
                    if (Documento.Contains("-"))
                    {
                        System.Windows.Forms.MessageBox.Show("Introduzca un número válido de documento/comprobante");
                    }
                    else
                    {
                        while (Documento.Length < 10)
                        {
                            Documento = "0" + Documento;
                        }
                        anticipos.anticiposopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocuAnt.Text, txtRUTAnt.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "Documento");
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }

                if (anticipos.ObjDatosAnticipos.Count > 0)
                {
                    GBDocsAPagar.Visibility = Visibility.Visible;
                    DGPagos.ItemsSource = null;
                    DGPagos.Items.Clear();

                    List<T_DOCUMENTOSAUX> partidaopen = new List<T_DOCUMENTOSAUX>();

                    for (int k = 0; k < anticipos.ObjDatosAnticipos.Count; k++)
                    {
                        T_DOCUMENTOSAUX partOpen = new T_DOCUMENTOSAUX();
                        partOpen.ISSELECTED = false;
                        partOpen.ACC = anticipos.ObjDatosAnticipos[k].ACC;
                        partOpen.CEBE = anticipos.ObjDatosAnticipos[k].CEBE;
                        partOpen.CLASE_CUENTA = anticipos.ObjDatosAnticipos[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = anticipos.ObjDatosAnticipos[k].CLASE_DOC;
                        partOpen.CME = anticipos.ObjDatosAnticipos[k].CME;
                        partOpen.COD_CLIENTE = anticipos.ObjDatosAnticipos[k].COD_CLIENTE;
                        partOpen.COND_PAGO = anticipos.ObjDatosAnticipos[k].COND_PAGO;
                        partOpen.CONTROL_CREDITO = anticipos.ObjDatosAnticipos[k].CONTROL_CREDITO;
                        partOpen.DIAS_ATRASO = anticipos.ObjDatosAnticipos[k].DIAS_ATRASO;
                        partOpen.ESTADO = anticipos.ObjDatosAnticipos[k].ESTADO;
                        partOpen.FECHA_DOC = anticipos.ObjDatosAnticipos[k].FECHA_DOC;
                        partOpen.FECVENCI = anticipos.ObjDatosAnticipos[k].FECVENCI;
                        partOpen.ICONO = anticipos.ObjDatosAnticipos[k].ICONO;
                        partOpen.MONEDA = anticipos.ObjDatosAnticipos[k].MONEDA;
                        partOpen.MONTO = anticipos.ObjDatosAnticipos[k].MONTO;
                        partOpen.MONTOF = anticipos.ObjDatosAnticipos[k].MONTOF;
                        partOpen.MONTOF_ABON = anticipos.ObjDatosAnticipos[k].MONTOF_ABON;
                        partOpen.MONTOF_PAGAR = anticipos.ObjDatosAnticipos[k].MONTOF_PAGAR;
                        partOpen.NDOCTO = anticipos.ObjDatosAnticipos[k].NDOCTO;
                        partOpen.NOMCLI = anticipos.ObjDatosAnticipos[k].NOMCLI;
                        partOpen.NREF = anticipos.ObjDatosAnticipos[k].NREF;
                        partOpen.RUTCLI = anticipos.ObjDatosAnticipos[k].RUTCLI;
                        partOpen.SOCIEDAD = anticipos.ObjDatosAnticipos[k].SOCIEDAD;
                        partidaopen.Add(partOpen);
                    }
                    DGPagos.ItemsSource = partidaopen;
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        private void limpiar()
        {
            GC.Collect();
        }
        //FUNCION QUE LIMPIA TODOS LOS ELEMENTOS PRESENTES EN LAS VIAS DE PAGO
        private void LimpiarViasDePago()
        {
            textBlock3.Text = "";
            textBlock4.Text = "";
            textBlock5.Text = "";
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
            DGCheque.ItemsSource = null;
            DGCheque.Items.Clear();
            DGMediosDePagos.ItemsSource = null;
            DGMediosDePagos.Items.Clear();
            cheques.Clear();
            txtNombreGira.Text = "";
            txtRUTGirador.Text = "";
            txtCodAut.Text = "";
            txtCodOp.Text = "";
            txtAsig.Text = "";
            txtSucursal.Text = "";
            txtPatente.Text = "";
            txtNumVenta.Text = "";
            txtObserv.Text = "";
            txtTasa.Text = "";
            txtTipoCamb.Text = "";
            txtNumDoc.Text = "";
            txtCantDoc.Text = "";
            DPFechVenc.Text = "";
            txtNumCuenta.Text = "";

            btnConfirPag.IsEnabled = false;

            btnAutorizacion.Visibility = Visibility.Collapsed;
            label46.Visibility = Visibility.Collapsed;
            cmbTipoTarjeta.Visibility = Visibility.Collapsed;
            label48.Visibility = Visibility.Collapsed;
            txtCodAut.Visibility = Visibility.Collapsed;
            label49.Visibility = Visibility.Collapsed;
            txtCodOp.Visibility = Visibility.Collapsed;
            label50.Visibility = Visibility.Collapsed;
            txtAsig.Visibility = Visibility.Collapsed;
            label32.Visibility = Visibility.Collapsed;
            DPFechActual.Visibility = Visibility.Collapsed;
            label25.Visibility = Visibility.Collapsed;
            DPFechVenc.Visibility = Visibility.Collapsed;
            label26.Visibility = Visibility.Collapsed;
            cmbBanco.Visibility = Visibility.Collapsed;
            label27.Visibility = Visibility.Collapsed;
            txtSucursal.Visibility = Visibility.Collapsed;
            label28.Visibility = Visibility.Collapsed;
            txtNumDoc.Visibility = Visibility.Collapsed;
            label30.Visibility = Visibility.Collapsed;
            txtNumCuenta.Visibility = Visibility.Collapsed;
            label31.Visibility = Visibility.Collapsed;
            txtRUTGirador.Visibility = Visibility.Collapsed;
            label38.Visibility = Visibility.Collapsed;
            txtNombreGira.Visibility = Visibility.Collapsed;
            label39.Content = "Número venta";
            label39.Visibility = Visibility.Collapsed;
            txtNumVenta.Visibility = Visibility.Collapsed;
            label40.Visibility = Visibility.Collapsed;
            txtObserv.Visibility = Visibility.Collapsed;
            label42.Visibility = Visibility.Collapsed;
            cmbIfinan.Visibility = Visibility.Collapsed;
            lblPatente.Visibility = Visibility.Collapsed;
            txtPatente.Visibility = Visibility.Collapsed;
            label33.Visibility = Visibility.Collapsed;
            txtCantDoc.Visibility = Visibility.Collapsed;
            label34.Visibility = Visibility.Collapsed;
            cmbIntervalo.Visibility = Visibility.Collapsed;

            GBViasPago.Visibility = Visibility.Collapsed;
            GC.Collect();
        }
        private void LimpiarEntradaDeViasDePago()
        {
            txtNombreGira.Text = "";
            txtRUTGirador.Text = "";
            txtCodAut.Text = "";
            txtCodOp.Text = "";
            txtAsig.Text = "";
            txtSucursal.Text = "";
            txtPatente.Text = "";
            txtNumVenta.Text = "";
            txtObserv.Text = "";
            txtTasa.Text = "";
            txtTipoCamb.Text = "";
            txtNumDoc.Text = "";
            txtCantDoc.Text = "";
            DPFechVenc.Text = "";
            txtNumCuenta.Text = "";
            cmbBancoProp.ItemsSource = null;
            cmbBancoProp.Items.Clear();
            cmbBanco.ItemsSource = null;
            cmbBanco.Items.Clear();
            cmbIfinan.ItemsSource = null;
            cmbIfinan.Items.Clear();
            cmbTipoTarjeta.ItemsSource = null;
            cmbTipoTarjeta.Items.Clear();
            cmbIntervalo.Text = "";

            GC.Collect();
        }
        #endregion
        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            bloquearCaja();
        }

        private void bloquearCaja()
        {
            timer.Stop();
            BloquearCaja bloquearcaja = new BloquearCaja();
            bloquearcaja.bloqueardesbloquearcaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, logApertura2);
            this.Close();
            GC.Collect();
        }

        private void chkDocFiscales_Checked(object sender, RoutedEventArgs e)
        {
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();

            DGDocDet.ItemsSource = null;
            DGDocCabec.Items.Clear();
            GC.Collect();
        }

        private void chkDocFiscales_Unchecked(object sender, RoutedEventArgs e)
        {
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocCabec.Items.Clear();
            GC.Collect();
        }
        private T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }

        private void txtNumDoc_TextChanged(object sender, TextChangedEventArgs e)
        {
            //Prueba
            bool digit = false;
            foreach (char value in txtNumDoc.Text)
            {
                digit = char.IsDigit(value);
            }

            if (digit)
            {
                ;
            }
            else
            {
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");
                txtNumDoc.Text = "";
            }
            GC.Collect();
        }

        private void LimpiarEntradasDeDatos()
        {
            txtRut.Text = "";
            txtDocu.Text = "";
            txtArchivo.Text = "";
            txtRUTAnt.Text = "";
            GC.Collect();
        }
        private void btnSalirCaja_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
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
        void ImpresionesDeDocumentosAutomaticas(string comprobante, string batch)
        {
            BusquedaReimpresiones busquedareimpresiones = new BusquedaReimpresiones();
            busquedareimpresiones.docsreimpresion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, comprobante, "", logApertura2[0].ID_REGISTRO, Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "X", Convert.ToString(lblSociedad.Content));
            if ((busquedareimpresiones.ViasPago.Count > 0) || (busquedareimpresiones.Documentos.Count > 0))
            {
                string InOut = "";

                if (busquedareimpresiones.Documentos[0].TIPO_DOCUMENTO == "N")
                    InOut = "Egreso";
                else
                    InOut = "Ingreso";

                ReimpresionComprobantes reimpresioncomprobantes = new ReimpresionComprobantes();
                reimpresioncomprobantes.reimprcomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                    , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, busquedareimpresiones.ViasPago
                    , busquedareimpresiones.Documentos);

                string Nombre = reimpresioncomprobantes.DatosCliente[0].NOMBRE;
                string RUT = reimpresioncomprobantes.DatosCliente[0].RUT;
                string Caja = reimpresioncomprobantes.DatosCaja[0].NOM_CAJA;
                string Documento = reimpresioncomprobantes.DatosDocumentos[0].NRO_DOCUMENTO;
                string Referencia = reimpresioncomprobantes.DatosCaja[0].ID_COMPROBANTE;
                string Pedido = reimpresioncomprobantes.DatosDocumentos[0].PEDIDO;
                string DocContable = reimpresioncomprobantes.NumDocCont;
                //LLAMADA AL FORM COMPROBANTES DESDE DONDE SE EMITE LA REIMPRESION DEL COMPROBANTE
                Comprobante frm = new Comprobante(reimpresioncomprobantes.DatosViaPago, reimpresioncomprobantes.DatosDocumentos, Nombre, RUT, Convert.ToString(textBlock7.Content)
                        , Convert.ToString(textBlock7.Content), Caja, Referencia, Documento, DocContable, InOut, logApertura2[0].MONEDA, Pedido, txtMandante.Text);
                if (reimpresioncomprobantes.DatosEmpresa.Count != 0)
                {
                    frm.txtSociedad.Text = reimpresioncomprobantes.DatosEmpresa[0].BUKRS;
                    frm.txtEmpresa.Text = reimpresioncomprobantes.DatosEmpresa[0].BUTXT;
                    frm.txtRIF.Text = reimpresioncomprobantes.DatosEmpresa[0].STCD1;
                }
                frm.Show();
            }
            GC.Collect();
        }

        private void LimpiarElementosDeCierreDeCaja()
        {
            txtCommCierre.Text = "";
            txtCommDif.Text = "";
        }
        private string Conversion(string montoingr)
        {
            string ViaPago = string.Empty;
            string resulCovert = string.Empty;

            string[] split = cmbVPMedioPag.SelectedValue.ToString().Split(new Char[] { '-' });
            ViaPago = split[0];

            if (ViaPago == "H " | ViaPago == "O ")
            {
                string USD = "USD";
                pagodocumentosingreso.Conversion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                        , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(ViaPago), MonedaCaja.Text, USD, montoingr, "0");
                resulCovert = string.Format(String.Format(CultureInfo.InvariantCulture,
                        "{0:0,0}", pagodocumentosingreso.ValorConvertido));
                resulCovert = Formato.FormatoMonedaChilena(resulCovert, "1");
                return resulCovert;
            }
            else
            {
                string EUR = "EUR ";
                pagodocumentosingreso.Conversion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                        , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(ViaPago), MonedaCaja.Text, EUR, montoingr, "0");
                resulCovert = string.Format("{0:0,0.##}", pagodocumentosingreso.ValorConvertido);
                return resulCovert;
            }
        }

        private string converOtrasMonedas(string montoingr)
        {
            string ViaPago = string.Empty;
            string resulCovert = string.Empty;

            string[] split = cmbVPMedioPag.SelectedValue.ToString().Split(new Char[] { '-' });
            ViaPago = split[0];

            if (ViaPago == "H " | ViaPago == "O ")
            {
                string USD = "USD";
                pagodocumentosingreso.Conversion2(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                        , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(ViaPago), USD, MonedaCaja.Text, "0", txtMontoFP.Text);
                resulCovert = Convert.ToString(pagodocumentosingreso.ValorConvertido);
                return resulCovert;
            }
            else
            {
                string EUR = "EUR ";
                pagodocumentosingreso.Conversion2(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                        , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(ViaPago), EUR, MonedaCaja.Text, "0", txtMontoFP.Text);
               resulCovert = Convert.ToString(pagodocumentosingreso.ValorConvertido);
                return resulCovert;
            }
        }

        private void ReportesCaja_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void txtMontoFP_PreviewKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            string ViaPago = string.Empty;
            string resulCovert = string.Empty;

            string[] split = cmbVPMedioPag.SelectedValue.ToString().Split(new Char[] { '-' });
            ViaPago = split[0];

            if (ViaPago == "H ")
            {
                string USD = "USD";
                pagodocumentosingreso.Conversion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                        , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(ViaPago), MonedaCaja.Text, USD, "0", Convert.ToString(textBlock4.Text));
                resulCovert = string.Format(String.Format(CultureInfo.InvariantCulture,
                        "{0:0,0}", pagodocumentosingreso.ValorConvertido));
                textBlock3.Text = string.Format(String.Format(CultureInfo.InvariantCulture,
                        "{0:0,0}", pagodocumentosingreso.ValorConvertido, resulCovert));
            }

        }

        private void btnValidarEfect_Click(object sender, RoutedEventArgs e)
        {
            string ViaPago = string.Empty;
            int MontoEfecCaja;
            double montovuelto;
            ViaPago = "E";

            pagodocumentosingreso.ValidarEfectivo(UsuarioCaja.Text, PassUserCaja.Text, txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, SociedCaja.Text, Convert.ToString(lblPais), idcaja.Text, UsuarioCaja.Text, ViaPago, logApertura2[0].ID_APERTURA);
            ValidEfec = pagodocumentosingreso.validar;
            double totviapago = Convert.ToDouble(textBlock3.Text);
            double totaPagar = Convert.ToDouble(textBlock4.Text);

            if (ValidEfec.Count()!= 0)
            {
                string ValConv = Convert.ToString(ValidEfec[0].MONTO);
                ValConv = ValConv.Replace(".", ",");
                double ValConv2 = Convert.ToDouble(ValConv);
                double ValConv3 = Math.Ceiling(ValConv2);
                MontoEfecCaja = Convert.ToInt32(ValConv3);
                string montovuel = Convert.ToString(txtVuelto.Text);
                montovuel = montovuel.Replace(",", "");
                montovuelto = Convert.ToDouble(montovuel);
                montovuelto = montovuelto * -1;

                if (MontoEfecCaja >= montovuelto)
                {
                    if (totviapago >= totaPagar)
                    {
                        btnConfirPag.IsEnabled = true;
                    }
                    else
                    {
                        btnConfirPag.IsEnabled = false;
                    }
                }
                else
                {
                    btnConfirPag.IsEnabled = false;
                    System.Windows.MessageBox.Show("Monto A Devolver excede Efectivo de Caja");
                }
            }
            else
            {
                btnConfirPag.IsEnabled = false;
                System.Windows.MessageBox.Show("Caja no cuenta con efectivo para devolver");
            }
        }
    }
}