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

namespace CajaIndigo.Vista.Vehiculos
{
    /// <summary>
    /// Interaction logic for Vehiculo.xaml
    /// </summary>
    public partial class Vehiculo : System.Windows.Window
    {
               private PdfTemplate totalPages;
        private PdfWriter Write;

        List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
        List<DetalleViasPago> cheques = new List<DetalleViasPago>();
        List<VIAS_PAGO_MASIVO> chequesMasiv = new List<VIAS_PAGO_MASIVO>();
        List<T_DOCUMENTOS> detalledocs = new List<T_DOCUMENTOS>();
        List<T_DOCUMENTOS> partidaseleccionadas = new List<T_DOCUMENTOS>();
        List<T_DOCUMENTOSAUX> partidaseleccionadasaux = new List<T_DOCUMENTOSAUX>();
        public List<ViasPago> ViasPagoTransaccion = new List<ViasPago>();
        //List<T_DOCUMENTOS> partidaselecc = new List<T_DOCUMENTOS>();
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

        Vista.PagoDocumento.PagoDocumento PagDocum;
        Vista.NotaCredito.NotaCredito NotaCredit;
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.Reportes.Reportes Reporte;

        List<LOG_APERTURA> logApertura = new List<LOG_APERTURA>();
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();
        FormatoMonedas Formato =  new  FormatoMonedas();


        public Vehiculo(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
        {
            try
            {
                InitializeComponent();
                int test = 0;
                GBInicio.Visibility = Visibility.Visible;
                GBMonitor.Visibility = Visibility.Visible;
                GBInicio.Visibility = Visibility.Collapsed;
                GBCommentCierre.Visibility = Visibility.Collapsed;
                textBlock6.Content = cajaconect;
                textBlock7.Content = usuariologg;
                textBlock8.Content = sucursal;
                textBlock9.Content = usuariotemp ;
                lblMonto.Content = Convert.ToString(monto);
                lblSociedad.Content = sociedad;

               // lblPais.Content = pais;
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
                //PaisCaja.Text = pais;
                DateTime result = DateTime.Today;
                datePicker1.Text = Convert.ToString(result);
                DGLogApertura.ItemsSource = null;
                DGLogApertura.Items.Clear();
                DGLogApertura.ItemsSource = logApertura;
                logApertura2 = logApertura;

                lblPais.Content = logApertura2[0].LAND;
                PaisCaja.Text = logApertura2[0].LAND;

                GC.Collect();  
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content),Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);               
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
                PagDocum = new PagoDocumento.PagoDocumento(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
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

        public Vehiculo()
        {
            InitializeComponent();
        }    
        private void bt_recaudacion(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }
        private void btnBuscarR_Click(object sender, RoutedEventArgs e)
        {
            listaRecaudacionVehiculo();
            GC.Collect();
        }
        private void btnSeleccionVehiculo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<IT_PAGOS> DocsASeleccionar = new List<IT_PAGOS>();
                List<AutorizacionViasPago> DocsAPagar = new List<AutorizacionViasPago>();
                List<IT_PAGOSAUX> partidaseleccionadasaux2 = new List<IT_PAGOSAUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGRecau.Items.Count > 0)
                {
                    for (int i = 0; i < DGRecau.Items.Count; i++)
                    {
                        if (i == 0)
                            DGRecau.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as IT_PAGOSAUX);
                        }
                        DGRecau.Items.MoveCurrentToNext();
                    }
                }

                List<string> listaTiposMoneda = new List<string>();

                for (int i = 0; i < partidaseleccionadasaux2.Count(); i++)
                {
                    if (!listaTiposMoneda.Contains(partidaseleccionadasaux2[i].WAERS) && partidaseleccionadasaux2[i].ISSELECTED == true)
                    {
                        listaTiposMoneda.Add(partidaseleccionadasaux2[i].WAERS);
                    }
                }         
                
                string MoneIni = string.Empty;
                bool Valida = true;

                MoneIni = listaTiposMoneda[0];
                int Cont = 0;

                foreach (String x in listaTiposMoneda)
                {
                    if(!MoneIni.Equals(x) && partidaseleccionadasaux2[Cont].ISSELECTED == true){
                        Valida = false;
                        break;
                    }
                }
                if (Valida == true)
                {
                    for (int i = 0; i <= partidaseleccionadasaux2.Count() - 1; i++)
                    {
                        if (partidaseleccionadasaux2[i].ISSELECTED == true)
                        {
                            IT_PAGOS partOpen = new IT_PAGOS();
                            partOpen.BANKN = partidaseleccionadasaux2[i].BANKN;
                            partOpen.CODBA = partidaseleccionadasaux2[i].CODBA;
                            partOpen.CODIN = partidaseleccionadasaux2[i].CODIN;
                            partOpen.CORRE = partidaseleccionadasaux2[i].CORRE;
                            partOpen.CTACE = partidaseleccionadasaux2[i].CTACE;
                            partOpen.CUOTA = partidaseleccionadasaux2[i].CUOTA;
                            partOpen.DBM_LICEXT = partidaseleccionadasaux2[i].DBM_LICEXT;
                            partOpen.DESCV = partidaseleccionadasaux2[i].DESCV;
                            partOpen.FEACT = partidaseleccionadasaux2[i].FEACT;
                            partOpen.FEVEN = partidaseleccionadasaux2[i].FEVEN;
                            partOpen.HKONT = partidaseleccionadasaux2[i].HKONT;
                            partOpen.INTER = partidaseleccionadasaux2[i].INTER;
                            partOpen.KKBER = partidaseleccionadasaux2[i].KKBER;
                            partOpen.KUNNR = partidaseleccionadasaux2[i].KUNNR;
                            partOpen.MINTE = partidaseleccionadasaux2[i].MINTE;
                            partOpen.MONTO = partidaseleccionadasaux2[i].MONTO;
                            partOpen.NOMBA = partidaseleccionadasaux2[i].NOMBA;
                            partOpen.NOMGI = partidaseleccionadasaux2[i].NOMGI;
                            partOpen.NOMIN = partidaseleccionadasaux2[i].NOMIN;
                            partOpen.NUDOC = partidaseleccionadasaux2[i].NUDOC;
                            partOpen.PRCTR = partidaseleccionadasaux2[i].PRCTR;
                            partOpen.RUTGI = partidaseleccionadasaux2[i].RUTGI;
                            partOpen.STAT = partidaseleccionadasaux2[i].STAT;
                            partOpen.STCD1 = partidaseleccionadasaux2[i].STCD1;
                            partOpen.TASAI = partidaseleccionadasaux2[i].TASAI;
                            partOpen.TOTIN = partidaseleccionadasaux2[i].TOTIN;
                            partOpen.VBELN = partidaseleccionadasaux2[i].VBELN;
                            partOpen.VIADP = partidaseleccionadasaux2[i].VIADP;
                            partOpen.WAERS = partidaseleccionadasaux2[i].WAERS;
                            DocsASeleccionar.Add(partOpen);
                        }
                    }
                  if (DocsASeleccionar.Count > 0)
                {
                    for (int j = 0; j < DocsASeleccionar.Count; j++)
                    {
                        AutorizacionViasPago Autoriza = new AutorizacionViasPago(DocsASeleccionar[j].VBELN, DocsASeleccionar[j].VIADP,DocsASeleccionar[j].DESCV, "", "", "", "", "");
                        DocsAPagar.Add(Autoriza);
                    }
                    bool Active = false;
                    int validador = 0;
                   
                    if (DocsAPagar.Count > 0)
                    {
                        Active = true;
                    }
                    for (int i = 0; i < DocsAPagar.Count; i++)
                    {
                        if ((DocsAPagar[i].VIADP == "S") | (DocsAPagar[i].VIADP == "R") | (DocsAPagar[i].VIADP == "G") | (DocsAPagar[i].VIADP == "F") | (DocsAPagar[i].VIADP == "B") | (DocsAPagar[i].VIADP == "U") | (DocsAPagar[i].VIADP == "P") | (DocsAPagar[i].VIADP == "L") | (DocsAPagar[i].VIADP == "K"))
                        {
                            validador = validador + 1;
                        }
                        else
                        {
                            validador = validador + 0;
                        }
                    }
                   
                    if (Active == true)
                    {
                        if (validador > 0)
                        {
                            GBAutorizacionVehiculos.Visibility = Visibility.Visible;
                        }
                        else
                        {
                            GBAutorizacionVehiculos.Visibility = Visibility.Visible;
                        }
                        DGAutorizacionVehiculos.ItemsSource = DocsAPagar;
                    }
            for (int i = 1; i < DGAutorizacionVehiculos.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGAutorizacionVehiculos.Items.MoveCurrentToFirst();
                }
                if (DGAutorizacionVehiculos.Items.CurrentItem != null)
                {
                }
                DGAutorizacionVehiculos.Items.MoveCurrentToNext();
                }
                btnPago.IsEnabled = true;
                GC.Collect();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Seleccione un registro para la recaudación de vehículos");
                }
                }
                else
                {
                    GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
                    DGAutorizacionVehiculos.ItemsSource = "";
                    System.Windows.Forms.MessageBox.Show("Los Pagos se deben realizar en la misma Moneda.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, ex.StackTrace);
                GC.Collect();
            }      
        }

        private void listaRecaudacionVehiculo()
        {
            string cadMensajes2 = string.Empty;
            string cadMensajes = string.Empty;
            string socied = string.Empty;
            DGRecau.ItemsSource = null;
            DGRecau.Items.Clear();
            DGAutorizacionVehiculos.ItemsSource = null;
            DGAutorizacionVehiculos.Items.Clear();
            String RUT = string.Empty;
            string RUTAux = "";
            foreach (char value in txtRuts.Text.ToUpper())
            {
                if (value != 8207)
                {
                    RUTAux = RUTAux + value;
                }
            }

            RUTAux = RUTAux.Trim();

            if (txtRuts.Text == "")
            {
                RUT = RUTAux;
            }
            else
            {
                RUT = DigitoVerificador(RUTAux);
            }

            if (RBRutRE.IsChecked == true)
            {
                if (RUT != RUTAux)
                {
                    System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                }
                else
                {
                    recauda.recauVehi(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocum.Text, RUT, Convert.ToString(lblSociedad.Content));
                    bapi_return2 = recauda.objReturn2;

                    for (int i = 0; i < bapi_return2.Count(); i++)
                    {
                        cadMensajes2 = cadMensajes2 + bapi_return2[i].MESSAGE + "<br>";
                        System.Windows.MessageBox.Show(cadMensajes2);
                    }
                }
            }
            else if (RBDocRE.IsChecked == true)
            {
                string Documento = "";
                Documento = txtDocum.Text;
                while (Documento.Length < 10)
                {
                    Documento = "0" + Documento;
                }
                txtDocum.Text = Documento;

                recauda.recauVehi(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocum.Text, RUT, Convert.ToString(lblSociedad.Content));
                bapi_return2 = recauda.objReturn2;

                for (int i = 0; i < bapi_return2.Count(); i++)
                {
                    cadMensajes2 = cadMensajes2 + bapi_return2[i].MESSAGE + "<br>";
                    System.Windows.MessageBox.Show(cadMensajes2);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
            }

            if (recauda.objPag.Count > 0)
            {
                GBDocs.Visibility = Visibility.Visible;
                btnActu.Visibility = Visibility.Collapsed;
                DGRecau.ItemsSource = null;
                DGRecau.Items.Clear();

                List<IT_PAGOSAUX> vehiculosseleccionados = new List<IT_PAGOSAUX>();
                for (int k = 0; k < recauda.objPag.Count; k++)
                {

                    IT_PAGOSAUX partOpen = new IT_PAGOSAUX();
                    partOpen.ISSELECTED = false;
                    partOpen.BANKN = recauda.objPag[k].BANKN;
                    partOpen.CODBA = recauda.objPag[k].CODBA;
                    partOpen.CODIN = recauda.objPag[k].CODIN;
                    partOpen.CORRE = recauda.objPag[k].CORRE;
                    partOpen.CTACE = recauda.objPag[k].CTACE;
                    partOpen.CUOTA = recauda.objPag[k].CUOTA;
                    partOpen.DBM_LICEXT = recauda.objPag[k].DBM_LICEXT;
                    partOpen.DESCV = recauda.objPag[k].DESCV;
                    partOpen.FEACT = recauda.objPag[k].FEACT;
                    partOpen.FEVEN = recauda.objPag[k].FEVEN;
                    partOpen.HKONT = recauda.objPag[k].HKONT;
                    partOpen.INTER = recauda.objPag[k].INTER;
                    partOpen.KKBER = recauda.objPag[k].KKBER;
                    partOpen.KUNNR = recauda.objPag[k].KUNNR;
                    partOpen.MINTE = recauda.objPag[k].MINTE;
                    partOpen.MONTO = recauda.objPag[k].MONTO;
                    partOpen.NOMBA = recauda.objPag[k].NOMBA;
                    partOpen.NOMGI = recauda.objPag[k].NOMGI;
                    partOpen.NOMIN = recauda.objPag[k].NOMIN;
                    partOpen.NUDOC = recauda.objPag[k].NUDOC;
                    partOpen.PRCTR = recauda.objPag[k].PRCTR;
                    partOpen.RUTGI = recauda.objPag[k].RUTGI;
                    partOpen.STAT = recauda.objPag[k].STAT;
                    partOpen.STCD1 = recauda.objPag[k].STCD1;
                    partOpen.TASAI = recauda.objPag[k].TASAI;
                    partOpen.TOTIN = recauda.objPag[k].TOTIN;
                    partOpen.VBELN = recauda.objPag[k].VBELN;
                    partOpen.VIADP = recauda.objPag[k].VIADP;
                    partOpen.WAERS = recauda.objPag[k].WAERS;
                    vehiculosseleccionados.Add(partOpen);
                }
                DGRecau.ItemsSource = vehiculosseleccionados;
                DGAutorizacionVehiculos.ItemsSource = null;
                DGAutorizacionVehiculos.Items.Clear();
            }
            GC.Collect();
        }

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
                                //}
                                //else
                                //{
                                //    System.Windows.MessageBox.Show("Ingrese el banco del depósito en cuenta corriente");
                                //}
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
                                                //if (txtAsig.Text != "")
                                                //{
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
                                                //}
                                                //else
                                                //{
                                                //    System.Windows.MessageBox.Show("Ingrese la asignación por tarjeta");
                                                //}
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
                                                //if (txtAsig.Text != "")
                                                //{
                                                //if (cmbTipoTarjeta.Text != "")
                                                //{
                                                IngresoFormasDePagoYMontos(MedioPago);
                                                //}
                                                //else
                                                //{
                                                //    System.Windows.MessageBox.Show("Ingrese el tipo de tarjeta");
                                                //}
                                                //}
                                                //else
                                                //{
                                                //    System.Windows.MessageBox.Show("Ingrese la asignación por tarjeta");
                                                //}
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
                                if (txtNumDoc.Text != "")
                                {
                                    if (txtMontoFP.Text != "")
                                    {
                                        if (txtCodAut.Text != "")
                                        {
                                            if (txtCodOp.Text != "")
                                            {
                                                //if (txtAsig.Text != "")
                                                //{
                                                //if (cmbTipoTarjeta.Text != "")
                                                //{
                                                IngresoFormasDePagoYMontos(MedioPago);
                                                //}
                                                //else
                                                //{
                                                //    System.Windows.MessageBox.Show("Ingrese el tipo de tarjeta");
                                                //}
                                                //}
                                                //else
                                                //{
                                                //    System.Windows.MessageBox.Show("Ingrese la asignación por tarjeta");
                                                //}
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
                        case "V": //Vale vista recibido
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

        //FUNCION QUE TOMA EL INGRESO DE DETALLES DE MEDIOS Y FORMAS DE PAGO, MONTOS Y LLENA EL GRID RESUMEN DE PAGOS (DGCHEQUES).
        private void IngresoFormasDePagoYMontos(string MedioPago)
        {
            try
            {
                String FechaVenct;
                if (DPFechActual.Text == "")
                {
                    DPFechActual.Text = datePicker1.Text;
                    DPFechActual.Text = Convert.ToString(datePicker1.SelectedDate.Value);
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
                string Valor = txtMontoFP.Text.Replace(".", "");
                Valor = Valor.Replace(",", ".");
                double monto = Convert.ToDouble(Valor);
                string moneda = cmbMoneda.Text;
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
                    detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, monto, moneda, banco, emisor
                        , num_cheque, cod_autorizacion, Convert.ToString(num_cuotas), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                        , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                        , cta_banco, ifinan, corre, zuonr, hkont, prctr, znop);
                    cheques.Add(detcheq);
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
                for (int i = 0; i <= items.Count - 1; i++)
                {
                    TotalVPagos = TotalVPagos + Convert.ToDouble(items[i].Monto);
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
                    textBlock3.Text = monedachil;
                    //decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = Formato.FormatoMoneda(Convert.ToString(MntTotalPend));
                    textBlock4.Text = monedachil2;
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = Formato.FormatoMoneda(Convert.ToString(ValorAux3));
                    if (monedachil3 == "00")
                    {
                        monedachil3 = "0";
                    }
                    textBlock5.Text = Convert.ToString(monedachil3);
                }
                else
                {
                    //decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = Formato.FormatoMonedaExtranjera(Convert.ToString(TotalVPagos));
                    textBlock3.Text = monedachil;
                    //decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = Formato.FormatoMonedaExtranjera(Convert.ToString(MntTotalPend));
                    textBlock4.Text = monedachil2;
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = Formato.FormatoMonedaExtranjera(Convert.ToString(ValorAux3));
                    textBlock5.Text = monedachil3;
                }
                if (Convert.ToDouble(textBlock5.Text) < 0)
                {
                    System.Windows.Forms.MessageBox.Show("Montos de vias de pago es superior a la cantidad a cancelar");
                    txtMontoFP.Text = "";
                    textBlock3.Text = "";
                    textBlock5.Text = "";
                    items.Clear();
                    cheques.Clear();
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

        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {
            GC.Collect();
        }

        //MUESTRA (VISIBILIDAD DE LOS ITEMS DE LOS MEDIOS DE PAGO A PARTIR DE LA SELECCION DEL COMBO DE MEDIOS DE PAGO
        private void cmbVPMedioPag_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
               // LimpiarEntradaDeViasDePago();
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
                                RFC_Combo_Bancos();
                                RFC_Carta_Curse();
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
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "F": //Cheque a fecha
                            {
                                RFC_Combo_Bancos();
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
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "G": //Cheque al día
                            {
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Visible;
                                //txtCodAuto.Visibility = Visibility.Visible;
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
                        case "M": //Contrato compra-venta
                            {
                                RFC_Combo_Bancos();
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
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "D": //Deposito a plazo
                            {
                                RFC_Combo_Bancos();
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
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "B": //Deposito en cliente corriente
                            {
                                //RFC BANCO PROPIO
                                RFC_Combo_BancosPropios();
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
                                break;
                            }
                        case "L": //Letras
                            {
                                RFC_Combo_Bancos();
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
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "P": //Pagaré
                            {
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Visible;
                                txtCantDoc.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Visible;
                                cmbIntervalo.Visibility = Visibility.Visible;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "E": //Pago en efectivo
                            {
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
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
                        case "S": //Tarjeta de crédito 
                            {

                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Visible;
                                label33.Content = "Número de cuotas";
                                txtCantDoc.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Visible;
                                cmbIntervalo.Visibility = Visibility.Visible;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;
                            }
                        case "R": //Tarjeta de débito
                            {
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
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
                        case "N": //Servipag
                            {
                                RFC_Combo_Tarjetas();
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
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
                                //RFC BANCO PROPIO
                                RFC_Combo_BancosPropios();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
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
                        case "V": //Vale vista recibido
                            {
                                RFC_Combo_Bancos();
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
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
                        case "A": //Vehiculo en parte de pago
                            {
                                RFC_Combo_Bancos();
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
                                //label39.Content = "Número VP";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Visible;
                                txtPatente.Visibility = Visibility.Visible;
                                label34.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
                                label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                label40.Margin = new Thickness(16, 64, 0, 0);
                                txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                txtObserv.Margin = new Thickness(136, 64, 0, 0);
                                txtObserv.Width = 389;
                                break;

                            }
                        case "1": //Documento Tributario
                            {
                                RFC_Combo_Bancos();
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
                                //label39.Content = "Número VP";
                                label39.Visibility = Visibility.Collapsed;
                                txtNumVenta.Visibility = Visibility.Collapsed;
                                label40.Visibility = Visibility.Collapsed;
                                txtObserv.Visibility = Visibility.Collapsed;
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
                                label42.Visibility = Visibility.Collapsed;
                                cmbIfinan.Visibility = Visibility.Collapsed;
                                lblPatente.Visibility = Visibility.Collapsed;
                                txtPatente.Visibility = Visibility.Collapsed;
                                label34.Visibility = Visibility.Collapsed;
                                txtCantDoc.Visibility = Visibility.Collapsed;
                                label33.Visibility = Visibility.Collapsed;
                                cmbIntervalo.Visibility = Visibility.Collapsed;
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

      private void btnConfirPag_Click(object sender, RoutedEventArgs e)
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
        }

                //RECAUDACION DE VEHICULOS
           private void btnPago_Click(object sender, RoutedEventArgs e)
                {
                    //DGRecau.SelectedItem
                    bool Variable01 = true;
                    List<AutorizacionViasPago> DocsAPagar = new List<AutorizacionViasPago>();
                    if (DGAutorizacionVehiculos.Items.Count > 0)
                    {
                        DGAutorizacionVehiculos.SelectAll();
                    }
                    Variable01 = false;

                    for (int i = 0; i < DGAutorizacionVehiculos.Items.Count; i++)
                    {
                        if (i == 0)
                        {
                            DGAutorizacionVehiculos.Items.MoveCurrentToFirst();
                        }
                        DocsAPagar.Add(DGAutorizacionVehiculos.Items.CurrentItem as AutorizacionViasPago);

                        DGAutorizacionVehiculos.Items.MoveCurrentToNext();
                    }

                    int validador = 0;
                    for (int j = 0; j < DocsAPagar.Count; j++)
                    {
                        if ((DocsAPagar[j].VIADP == "F") | (DocsAPagar[j].VIADP == "G"))
                        {
                            if (DocsAPagar[j].NUMTARJETA != "")
                            {

                                if (DocsAPagar[j].AUTORIZACION != "")
                                {
                                    Variable01 = true;
                                }
                                else
                                {
                                    Variable01 = false;
                                    validador = validador + 1;
                                    System.Windows.Forms.MessageBox.Show("Ingrese el código de autorización del cheque  N° " + DocsAPagar[j].VBELN);
                                }
                            }
                            else
                            {
                                validador = validador + 1;
                                System.Windows.Forms.MessageBox.Show("Ingrese el número de cheque en documento N° " + DocsAPagar[j].VBELN);
                            }
                        }

                        if ((DocsAPagar[j].VIADP == "P") | (DocsAPagar[j].VIADP == "L") | (DocsAPagar[j].VIADP == "K"))
                        {

                            if (DocsAPagar[j].NUMTARJETA != "")
                            {
                                Variable01 = true;
                            }
                            else
                            {
                                Variable01 = false;
                                validador = validador + 1;
                                System.Windows.Forms.MessageBox.Show("Ingrese el número de documento de vía de pago en documento N°" + DocsAPagar[j].VBELN);
                            }
                        }
                        if ((DocsAPagar[j].VIADP == "U") | (DocsAPagar[j].VIADP == "B"))
                        {

                            if (DocsAPagar[j].FEC_EMISION != "")
                            {
                                Variable01 = true;
                            }
                            else
                            {
                                Variable01 = false;
                                validador = validador + 1;
                                System.Windows.Forms.MessageBox.Show("Ingrese la fecha de emisión del documento " + DocsAPagar[j].VBELN);
                            }
                        }
                        if ((DocsAPagar[j].VIADP == "S") | (DocsAPagar[j].VIADP == "R"))
                        {
                            if (DocsAPagar[j].NUMTARJETA != "")
                            {
                                if (DocsAPagar[j].OPERACION != "")
                                {
                                    if (DocsAPagar[j].AUTORIZACION != "") 
                                    {
                                        Variable01 = true;
                                    }
                                    else
                                    {
                                        Variable01 = false;
                                        validador = validador + 1;
                                        System.Windows.Forms.MessageBox.Show("Ingrese el código de autorización  en documento N° " + DocsAPagar[j].VBELN);
                                    }
                                }
                                else
                                {
                                    validador = validador + 1;
                                    System.Windows.Forms.MessageBox.Show("Ingrese el código de operación en documento N° " + DocsAPagar[j].VBELN);
                                }
                            }
                            else
                            {
                                validador = validador + 1;
                                System.Windows.Forms.MessageBox.Show("Ingrese el número de tarjeta crédito ó débito en documento N° " + DocsAPagar[j].VBELN);
                            }

                        }
                        if ((DocsAPagar[j].VIADP != "S") | (DocsAPagar[j].VIADP != "R") | (DocsAPagar[j].VIADP != "F") | (DocsAPagar[j].VIADP != "G") | (DocsAPagar[j].VIADP != "U") | (DocsAPagar[j].VIADP != "B"))
                        {
                            validador = validador + 0;
                        }
                    }
                    if (validador < 1)
                    {
                        RecaudacionVehiculos(DocsAPagar, logApertura2[0].ID_REGISTRO);
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("No se pudo procesar la recaudación de vehículos");
                    }
                    GC.Collect();
                }

                //BOTON QUE LLEVA EL DATO SELECCIONADO DESDE EL MONITOR 
                private void btnPagoMonitor_Click(object sender, RoutedEventArgs e)
                {
                    //LimpiarViasDePago();
                    GBDocsAPagar.Visibility = Visibility.Collapsed;
                    GBViasPago.Visibility = Visibility.Collapsed;
                    ListaDocumentosPendientesDesdeMonitor();
                    GBDocsAPagar.Visibility = Visibility.Visible;
                    GBViasPago.Visibility = Visibility.Visible;
                    txtDocum.Text = "";
                    GC.Collect();
                }

                //FUNCION QUE TRAE LOS DOCUMENTOS DESDE EL MONITOR
                private void ListaDocumentosPendientesDesdeMonitor()
                {
                    try
                    {
                        timer.Stop();
                        for (int i = detalledocs.Count - 1; i >= 0; --i)
                        {
                            detalledocs.RemoveAt(i);
                        }
                        DGPagos.ItemsSource = null;
                        DGPagos.Items.Clear();
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
                        //Calculo del monto para los documentos y partidas abiertas seleccionadas.
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
                        //List<T_DOCUMENTOS> partidaopen = new List<T_DOCUMENTOS>;
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
                            //RFC para consulta de estatus de cobro del cliente selecccionado
                            DateTime Anual = datePicker1.SelectedDate.Value;
                            String EjercicioValue = Convert.ToString(Anual.Year);
                            //RFC que retorna las formas de pago de acuerdo a los registros seleccionados
                            MatrizDePago matrizpago = new MatrizDePago();
                            // List<ViasPago> LVP = new List<ViasPago>();
                            String Protesto = "";
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
                                }
                                else
                                {
                                    //string moneda = string.Format("{0:0,0.##}", Monto);
                                    textBlock4.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(Monto));
                                    txtMontoFP.Text = Formato.FormatoMonedaExtranjera(Convert.ToString(Monto));
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

                //CONEXION A LA RFC DEL MONITOR EN MODO MANUAL
                private void btnRefresh_Click(object sender, RoutedEventArgs e)
                {
                    try
                    {
                        //RFC del Monitor por boton Refresh 
                        monitor.ObjDatosMonitor.Clear();
                        monitor.monitor(Convert.ToString(datePicker1.SelectedDate.Value), Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblSociedad.Content));
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

                //CLICK QUE MUESTRA EL FORM DE RESUMEN DE PAGOS CON TODA LA INFORMACION DE LAS VIAS DE PAGO 
                private void btnResPagos_Click(object sender, RoutedEventArgs e)
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

                private void ChkVehiculos_Checked(object sender, RoutedEventArgs e)
                {
                    List<IT_PAGOSAUX> partidaseleccionadasaux2 = new List<IT_PAGOSAUX>();
                    partidaseleccionadasaux2.Clear();
                    if (this.DGRecau.Items.Count > 0)
                    {
                        for (int i = 0; i < DGRecau.Items.Count; i++)
                        {
                            if (i == 0)
                                DGRecau.Items.MoveCurrentToFirst();
                            {
                                partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as IT_PAGOSAUX);
                            }
                            DGRecau.Items.MoveCurrentToNext();
                        }
                    }
                    for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                    {
                        partidaseleccionadasaux2[k].ISSELECTED = true;
                    }
                    DGRecau.ItemsSource = null;
                    DGRecau.Items.Clear();
                    DGRecau.ItemsSource = partidaseleccionadasaux2;
                    GC.Collect();
                }

                private void ChkVehiculos_UnChecked(object sender, RoutedEventArgs e)
                {
                    List<IT_PAGOSAUX> partidaseleccionadasaux2 = new List<IT_PAGOSAUX>();
                    partidaseleccionadasaux2.Clear();
                    if (this.DGRecau.Items.Count > 0)
                    {
                        for (int i = 0; i < DGRecau.Items.Count; i++)
                        {
                            if (i == 0)
                                DGRecau.Items.MoveCurrentToFirst();
                            {
                                partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as IT_PAGOSAUX);
                            }
                            DGRecau.Items.MoveCurrentToNext();
                        }
                    }
                    for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                    {
                        partidaseleccionadasaux2[k].ISSELECTED = false;
                    }
                    DGRecau.ItemsSource = null;
                    DGRecau.Items.Clear();
                    DGRecau.ItemsSource = partidaseleccionadasaux2;
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

                private void DGMonitor_SelectionChanged(object sender, SelectionChangedEventArgs e)
                {

                    timer.Stop();
                    GC.Collect();
                }

                private void RBDocRE_Checked(object sender, RoutedEventArgs e)
                {
                    txtRuts.Text = "";
                    txtRuts.Visibility = Visibility.Collapsed;
                    txtDocum.Text = "";
                    txtDocum.Visibility = Visibility.Visible;
                    GBDocsAPagar.Visibility = Visibility.Collapsed;
                    GBViasPago.Visibility = Visibility.Collapsed;
                  //Limpiar aqui el resumen de las vias de pago
                    LimpiarViasDePago();
                    GC.Collect();
                }
                private void RBRutRE_Checked(object sender, RoutedEventArgs e)
                {
                    txtRuts.Text = "";
                    txtRuts.Visibility = Visibility.Visible;
                    txtDocum.Text = "";
                    txtDocum.Visibility = Visibility.Collapsed;
                    GBDocsAPagar.Visibility = Visibility.Collapsed;
                    GBViasPago.Visibility = Visibility.Collapsed;    
                  //Limpiar aqui el resumen de las vias de pago
                    LimpiarViasDePago();
                    GC.Collect();
                }

                //FUNCION QUE LIMPIA TODOS LOS ELEMENTOS PRESENTES EN LAS VIAS DE PAGO
                private void LimpiarViasDePago()
                {
                    //Limpiar aqui el resumen de las vias de pago
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
                    //label41.Visibility = Visibility.Collapsed;
                    //txtCodAuto.Visibility = Visibility.Collapsed;
                    label42.Visibility = Visibility.Collapsed;
                    cmbIfinan.Visibility = Visibility.Collapsed;
                    lblPatente.Visibility = Visibility.Collapsed;
                    txtPatente.Visibility = Visibility.Collapsed;
                    label33.Visibility = Visibility.Collapsed;
                    txtCantDoc.Visibility = Visibility.Collapsed;
                    label34.Visibility = Visibility.Collapsed;
                    cmbIntervalo.Visibility = Visibility.Collapsed;
                    GBViasPago.Visibility = Visibility.Collapsed;
                    GBDocs.Visibility = Visibility.Collapsed;
                    GC.Collect();
                }

                private void txtNumDoc_TextChanged(object sender, TextChangedEventArgs e)
                {
                    bool digit = true;
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
                private void RecaudacionVehiculos(List<AutorizacionViasPago> Autorizacion, string Id_Apertura)
                {
                    try
                    {
                        List<VIAS_PAGO_VEHI> lisParamt = new List<VIAS_PAGO_VEHI>();
                        List<DOCUMENTO_CAB> lisParamtCab = new List<DOCUMENTO_CAB>();
                        List<ACT_FPAGOS> lisParamPag = new List<ACT_FPAGOS>();
                        List<IT_PAGOS_CAB> lisItCab = new List<IT_PAGOS_CAB>();
                        VIAS_PAGO_VEHI paramt = new VIAS_PAGO_VEHI();
                        String NCajero = Convert.ToString(textBlock7.Content);
                        string RutClie = Convert.ToString(txtRuts.Text);
                        string Moneda = Convert.ToString(cmbMoneda.SelectedItem);
                        string Pais = Convert.ToString(lblPais.Content);


                        lbTitulo.Visibility = Visibility.Visible;
                        GBViasPago.Visibility = Visibility.Collapsed;
                        //Calculo del monto para los documentos y partidas abiertas seleccionadas.
                        viapago.Clear();
                        List<IT_PAGOSAUX> partidaseleccionadasaux2 = new List<IT_PAGOSAUX>();
                        partidaseleccionadasaux2.Clear();
                        if (this.DGRecau.Items.Count > 0)
                        {
                            for (int i = 0; i < DGRecau.Items.Count; i++)
                            {
                                if (i == 0)
                                    DGRecau.Items.MoveCurrentToFirst();
                                {
                                    partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as IT_PAGOSAUX);
                                }
                                DGRecau.Items.MoveCurrentToNext();
                            }
                        }
                        for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                        {
                            if (partidaseleccionadasaux2[k].ISSELECTED == true)
                            {
                                IT_PAGOS partOpen = new IT_PAGOS();
                                partOpen.BANKN = partidaseleccionadasaux2[k].BANKN;
                                partOpen.CODBA = partidaseleccionadasaux2[k].CODBA;
                                partOpen.CODIN = partidaseleccionadasaux2[k].CODIN;
                                partOpen.CORRE = partidaseleccionadasaux2[k].CORRE;
                                partOpen.CTACE = partidaseleccionadasaux2[k].CTACE;
                                partOpen.CUOTA = partidaseleccionadasaux2[k].CUOTA;
                                partOpen.DBM_LICEXT = partidaseleccionadasaux2[k].DBM_LICEXT;
                                partOpen.DESCV = partidaseleccionadasaux2[k].DESCV;
                                partOpen.FEACT = partidaseleccionadasaux2[k].FEACT;
                                partOpen.FEVEN = partidaseleccionadasaux2[k].FEVEN;
                                partOpen.HKONT = partidaseleccionadasaux2[k].HKONT;
                                partOpen.INTER = partidaseleccionadasaux2[k].INTER;
                                partOpen.KKBER = partidaseleccionadasaux2[k].KKBER;
                                partOpen.KUNNR = partidaseleccionadasaux2[k].KUNNR;
                                partOpen.MINTE = partidaseleccionadasaux2[k].MINTE;
                                partOpen.MONTO = partidaseleccionadasaux2[k].MONTO;
                                partOpen.NOMBA = partidaseleccionadasaux2[k].NOMBA;
                                partOpen.NOMGI = partidaseleccionadasaux2[k].NOMGI;
                                partOpen.NOMIN = partidaseleccionadasaux2[k].NOMIN;
                                partOpen.NUDOC = partidaseleccionadasaux2[k].NUDOC;
                                partOpen.PRCTR = partidaseleccionadasaux2[k].PRCTR;
                                partOpen.RUTGI = partidaseleccionadasaux2[k].RUTGI;
                                partOpen.STAT = partidaseleccionadasaux2[k].STAT;
                                partOpen.STCD1 = partidaseleccionadasaux2[k].STCD1;
                                partOpen.TASAI = partidaseleccionadasaux2[k].TASAI;
                                partOpen.TOTIN = partidaseleccionadasaux2[k].TOTIN;
                                partOpen.VBELN = partidaseleccionadasaux2[k].VBELN;
                                partOpen.VIADP = partidaseleccionadasaux2[k].VIADP;
                                partOpen.WAERS = partidaseleccionadasaux2[k].WAERS;
                                viapago.Add(partOpen);
                            }
                        }
                        if (viapago.Count > 0)
                        {
                            //pasar datos al estrucutra principal
                            VIAS_PAGO_VEHI ls_paramt;
                            decimal total = 0;

                            for (int j = 0; j < viapago.Count; j++)
                            {
                                ls_paramt = new VIAS_PAGO_VEHI();

                                if (viapago[j].VIADP == "F")
                                {
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.BANCO = viapago[j].CODBA;
                                    ls_paramt.NUM_CHEQUE = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.NUM_CUENTA = viapago[j].CTACE;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.COD_AUTORIZACION = Autorizacion[j].AUTORIZACION;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                if (viapago[j].VIADP == "C")
                                {
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.BANCO = viapago[j].CODBA;
                                    ls_paramt.NUM_CHEQUE = viapago[j].NUDOC;
                                    ls_paramt.FECHA_VENC = viapago[j].FEACT;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.NUM_CUENTA = viapago[j].CTACE;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }

                                if (viapago[j].VIADP == "E")
                                {
                                    ls_paramt.TEXTO_POSICION = "";
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                if (viapago[j].VIADP == "H")
                                {
                                    ls_paramt.TEXTO_POSICION = "";
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }

                                //Efectivo Euro 09.09.2015
                                if (viapago[j].VIADP == "J")
                                {
                                    ls_paramt.TEXTO_POSICION = "";
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }

                                // TRANSFERENCIA ELECTRONICA
                                if (viapago[j].VIADP == "O")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].RUTGI;

                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.HKONT = viapago[j].HKONT;
                                    ls_paramt.FECHA_EMISION = Autorizacion[j].FEC_EMISION;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }


                                if (viapago[j].VIADP == "D")
                                {
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.BANCO = viapago[j].CODBA;
                                    ls_paramt.NUM_CHEQUE = viapago[j].NUDOC;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.NUM_CUENTA = viapago[j].CTACE;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // Tarjeta Credito
                                if (viapago[j].VIADP == "S")
                                {
                                    ls_paramt.NUM_TARJETA = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.NUM_CUOTAS = "001";
                                    ls_paramt.COD_AUTORIZACION = Autorizacion[j].AUTORIZACION;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    //ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    //ls_paramt.ZUONR = Autorizacion[j].ASIGNACION;
                                    ls_paramt.ZNOP = Autorizacion[j].OPERACION;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // tarjeta Debito
                                if (viapago[j].VIADP == "R")
                                {
                                    ls_paramt.NUM_TARJETA = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.NUM_CUOTAS = "001";
                                    ls_paramt.COD_AUTORIZACION = Autorizacion[j].AUTORIZACION;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    //ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    //ls_paramt.ZUONR = Autorizacion[j].ASIGNACION;
                                    ls_paramt.ZNOP = Autorizacion[j].OPERACION;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                if (viapago[j].VIADP == "N")
                                {
                                    ls_paramt.NUM_TARJETA = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.NUM_CUOTAS = "001";
                                    ls_paramt.COD_AUTORIZACION = Autorizacion[j].AUTORIZACION;
                                  //ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    //ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    //ls_paramt.ZUONR = Autorizacion[j].ASIGNACION;
                                    ls_paramt.ZNOP = Autorizacion[j].OPERACION;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                        //VEHICULO EN PARTE DE PAGO
                        if (viapago[j].VIADP == "A")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].RUTGI;
                                    ls_paramt.PATENTE = viapago[j].DBM_LICEXT;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.NUM_VENTA = viapago[j].VBELN + "-" + "Vehiculo en Parte Pago";
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                if (viapago[j].VIADP == "K")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    ls_paramt.IFINAN = viapago[j].KUNNR;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.NUM_VENTA = viapago[j].VBELN;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.CARTA_CURSE = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                if (viapago[j].VIADP == "L")
                                {
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.NUM_CUOTAS = "001";
                                    ls_paramt.PAGARE = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                //PAGARE
                                if (viapago[j].VIADP == "P")
                                {
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.NUM_CUOTAS = "001";
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.PAGARE = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // DEPOSITO EN CUENTA CORRIENTE
                                if (viapago[j].VIADP == "B")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].STCD1;
                                    //ls_paramt.BANCO = viapago[j].CODBA;
                                    //ls_paramt.NUM_CUENTA = viapago[j].CTACE;
                                    ls_paramt.HKONT = viapago[j].HKONT;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.NUM_DEPOSITO = viapago[j].NUDOC;
                                    ls_paramt.FECHA_EMISION = Autorizacion[j].FEC_EMISION;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // CHEQUE AL DIA
                                if (viapago[j].VIADP == "G")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].RUTGI;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.BANCO = viapago[j].CODBA;
                                    ls_paramt.NUM_CHEQUE = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.FECHA_VENC = viapago[j].FEACT;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.NUM_CUENTA = viapago[j].CTACE;
                                    ls_paramt.COD_AUTORIZACION = Autorizacion[j].AUTORIZACION;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // VALE VISTA
                                if (viapago[j].VIADP == "V")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].RUTGI;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.NUM_CHEQUE = viapago[j].NUDOC;
                                    ls_paramt.BANCO = viapago[j].CODBA;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEVEN;
                                    ls_paramt.NUM_CUENTA = viapago[j].CTACE;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.NUM_VALE_VISTA = Autorizacion[j].NUMTARJETA;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // TRANSFERENCIA ELECTRONICA
                                if (viapago[j].VIADP == "U")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].RUTGI;

                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.FECHA_VENC = viapago[j].FEVEN;
                                    ls_paramt.HKONT = viapago[j].HKONT;
                                    ls_paramt.FECHA_EMISION = Autorizacion[j].FEC_EMISION;
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.PRCTR = viapago[j].PRCTR;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.SUCURSAL = "999";
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }
                                // CONTRATO COMPRAVENTA
                                if (viapago[j].VIADP == "M")
                                {
                                    string rutclie = Convert.ToString(txtRuts.Text);
                                    rutclie = viapago[j].RUTGI;
                                    ls_paramt.PATENTE = viapago[j].DBM_LICEXT;
                                    ls_paramt.NOMBRE_GIRADOR = viapago[j].NOMGI;
                                    ls_paramt.EMISOR = viapago[j].RUTGI;
                                    ls_paramt.HKONT = viapago[j].HKONT;
                                    ls_paramt.FECHA_EMISION = viapago[j].FEACT;
                                    ls_paramt.MONTO = viapago[j].MONTO;
                                    ls_paramt.NUM_VENTA = viapago[j].VBELN + "-" + "Contrato de Compraventa";
                                    ls_paramt.LAND = Pais;
                                    ls_paramt.MONEDA = viapago[j].WAERS;
                                    ls_paramt.VIA_PAGO = viapago[j].VIADP;
                                    ls_paramt.NUM_VALE_VISTA = viapago[j].NUDOC;
                                    ls_paramt.ZUONR = viapago[j].VBELN;
                                    ls_paramt.CORRE = viapago[j].CORRE;
                                    lisParamt.Add(ls_paramt);
                                }

                                string valor = viapago[j].MONTO;
                                valor = valor.Replace(".", "");
                                decimal valor2 = Convert.ToDecimal(valor);
                                total = total + valor2;

                                //Cambia estatus a documento 
                                ACT_FPAGOS ls_paramtPag;
                                ls_paramtPag = new ACT_FPAGOS();
                                ls_paramtPag.CORRE = viapago[j].CORRE;
                                ls_paramtPag.VBELN = viapago[j].VBELN;
                                lisParamPag.Add(ls_paramtPag);
                                // Pasar Estructura Cabecera
                                DOCUMENTO_CAB ls_paramtCab;
                                ls_paramtCab = new DOCUMENTO_CAB();
                                ls_paramtCab.SOCIEDAD = Convert.ToString(lblSociedad.Content);
                                ls_paramtCab.CAJERO_RESP = NCajero;
                                ls_paramtCab.LAND = Pais;
                                ls_paramtCab.CLASE_DOC = "DZ";
                                ls_paramtCab.MONEDA = viapago[0].WAERS;
                                ls_paramtCab.TIPO_DOCUMENTO = viapago[0].VIADP;
                                if (RutClie == "")
                                {
                                    ls_paramtCab.CLIENTE = viapago[j].STCD1;
                                }
                                else
                                {
                                    ls_paramtCab.CLIENTE = RutClie;
                                }
                                if ((ls_paramtCab.ACC == "") | (ls_paramtCab.ACC == null))
                                {
                                    ls_paramtCab.ACC = viapago[j].KKBER;
                                }
                                ls_paramtCab.NRO_DOCUMENTO = viapago[j].NUDOC;
                                ls_paramtCab.CEBE = viapago[j].PRCTR;
                                ls_paramtCab.ID_APERTURA = Id_Apertura;
                                lisParamtCab.Add(ls_paramtCab);
                            }

                            RECAUDA.PagaVehicu(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), "0000", lisParamt, lisParamtCab, lisParamPag, viapago[0].WAERS, RutClie, Convert.ToString(lblSociedad.Content), "", Convert.ToString(total), Convert.ToString(lblPais.Content));

                            bapi_return2 = recauda.objReturn2;

                            for (int i = 0; i < bapi_return2.Count(); i++)
                            {
                                string cadMensajes = "";
                                cadMensajes = cadMensajes + bapi_return2[i].MESSAGE + "<br>";
                                System.Windows.MessageBox.Show(cadMensajes);
                            }
                            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
                            DGAutorizacionVehiculos.ItemsSource = null;
                            DGAutorizacionVehiculos.Items.Clear();
                            btnPago.IsEnabled = false;

                            string Mensaje = "";
                            string Mensaje1 = "Pedido DBM";
                            string Mensaje2 = "Doc Financiero";
                            string Mensaje3 = "Comprobante Ingreso";
                            string numeroContacto = RECAUDA.DOCUMENTO;
                            string numeroComprobante = RECAUDA.DOCUMENTO2;

                            logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);
                            listaRecaudacionVehiculo();

                            if (RECAUDA.errormessage != "")
                            {
                                System.Windows.MessageBox.Show(RECAUDA.errormessage);
                            }

                            if (numeroComprobante != "0000000000")
                            {
                                Mensaje = "Pedido DBM:" + viapago[0].VBELN + "\n" + "DocFinan:" + numeroContacto + "\n" + "NumComproba:" + numeroComprobante + "\n";
                                System.Windows.MessageBox.Show(Mensaje);
                               ImpresionesDeDocumentosAutomaticas(numeroComprobante, "X");
                            }
                            else
                            {
                                System.Windows.MessageBox.Show("No se generó comprobante de pago");
                            }
                        }
                        GC.Collect();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        GC.Collect();
                    }
                }
                private void RFC_Carta_Curse()
                {
                    maestrofinanc.maestroifinan(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), "K", Convert.ToString(lblSociedad.Content));
                    if (maestrofinanc.T_Retorno.Count > 0)
                    {

                        cmbIfinan.ItemsSource = null;
                        cmbIfinan.Items.Clear();


                        List<string> listabancos = new List<string>();
                        listabancos.Clear();
                        for (int i = 0; i < maestrofinanc.T_Retorno.Count; i++)
                        {
                            listabancos.Add(maestrofinanc.T_Retorno[i].KUNNR + " - " + maestrofinanc.T_Retorno[i].MCOD1);
                        }
                        // Llena Combobox Pagos Masivos
                        cmbIfinan.ItemsSource = listabancos;
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

                        List<string> listabancos = new List<string>();
                        listabancos.Clear();

                        for (int i = 0; i < maestrobancos.T_Retorno.Count; i++)
                        {
                            listabancos.Add(maestrobancos.T_Retorno[i].BANKL + " - " + maestrobancos.T_Retorno[i].BANKA);
                        }
                        cmbBanco.ItemsSource = listabancos;
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
                    maestrobancos.maestrobancos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbMoneda.Text), Convert.ToString(lblSociedad.Content));
                    if (maestrobancos.T_Retorno2.Count > 0)
                    {
                        cmbBancoProp.ItemsSource = null;
                        cmbBancoProp.Items.Clear();
                        cmbCuentasBancosProp.ItemsSource = null;
                        cmbCuentasBancosProp.Items.Clear();
                        cmbBanco.ItemsSource = null;
                        cmbBanco.Items.Clear();

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
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("No existen datos de bancos propios en el sistema");
                    }
                    GC.Collect();
                }

                //RFC QUE LLENA LAS TARJETAS EN MEDIOS DE PAGOS
                private void RFC_Combo_Tarjetas()
                {
                    cmbTipoTarjeta.ItemsSource = null;
                    cmbTipoTarjeta.Items.Clear();

                    if (cmbVPMedioPag.Text == "")
                    {
                        maestrotarjetas.maestrotarjetas(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbVPMedioPag.Text));
                    }
                    
                    if (maestrotarjetas.T_Retorno.Count > 0)
                    {
                        cmbTipoTarjeta.ItemsSource = null;
                        cmbTipoTarjeta.Items.Clear();

                        List<string> listatarjetas = new List<string>();
                        listatarjetas.Clear();
                        for (int i = 0; i < maestrotarjetas.T_Retorno.Count; i++)
                        {
                            listatarjetas.Add(maestrotarjetas.T_Retorno[i].CCINS + " - " + maestrotarjetas.T_Retorno[i].VTEXT);
                        }
                        cmbTipoTarjeta.ItemsSource = listatarjetas;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("No existen datos de " + Convert.ToString(cmbVPMedioPag.Text).Substring(3, Convert.ToString(cmbVPMedioPag.Text).Length - 3) + " en el sistema");
                    }
                    GC.Collect();
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

                private void CierreCaja_Click(object sender, RoutedEventArgs e)
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
                    GC.Collect();
                }

        private void ReportesCaja_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }
    }
        }
