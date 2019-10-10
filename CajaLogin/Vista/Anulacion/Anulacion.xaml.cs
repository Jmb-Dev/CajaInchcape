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
using CajaIndigo.AppPersistencia.Class.AnulacionComprobantes.Estructura;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;
using CajaIndigo;
using System.Text.RegularExpressions;
using System.Reflection;

namespace CajaIndigo.Vista.Anulacion
{
    /// <summary>
    /// Interaction logic for Anulacion.xaml
    /// </summary>
    public partial class Anulacion : System.Windows.Window
    {
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
        string Accion = string.Empty;
        string ViaPago = string.Empty;
        double montoEfecCaja;
        double montoDocu;
        DateTime fechadia ;

        Vista.PagoDocumento.PagoDocumento PagDocum;
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.NotaCredito.NotaCredito NotaCredit;
        Vista.Reportes.Reportes Reporte;

        List<LOG_APERTURA> logApertura = new List<LOG_APERTURA>();
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();

        string IdComprobante = "";

        public Anulacion()
        {
            InitializeComponent();
        }

        public Anulacion(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
        {
            try
            {
                InitializeComponent();

                List<string> myItemsCollection = new List<string>();
                myItemsCollection.Add(moneda);
                GBInicio.Visibility = Visibility.Collapsed;
                textBlock6.Content = cajaconect;
                textBlock7.Content = usuariologg;
                textBlock8.Content = sucursal;
                textBlock9.Content = usuariotemp ;  
                lblMonto.Content = Convert.ToString(monto);
                lblSociedad.Content = sociedad;
                txtIdSistema.Text = IdSistema;
                txtInstancia.Text = Instancia;
                txtMandante.Text = mandante;
                txtSapRouter.Text = SapRouter;
                txtServer.Text = server;
                txtIdioma.Text = idioma;
                UsuarioCaja.Text = usuariologg;
                txtUserAnula.Text = usuariologg;
                PassUserCaja.Text = passlogg;
                idcaja.Text = cajaconect;
                NomCaja.Text = sucursal;
                SociedCaja.Text = sociedad;
                MonedaCaja.Text = moneda;
                PaisCaja.Text = pais;
                        
                //lblPais.Content = pais;
                lblPassword.Content = passlogg;
                DateTime result = DateTime.Today;
                calendario.Text = Convert.ToString(result);
                DGLogApertura.ItemsSource = null;
                DGLogApertura.Items.Clear();
                DGLogApertura.ItemsSource = logApertura;
                logApertura2 = logApertura;
                lblPais.Content = logApertura2[0].LAND;
                txtpais.Text = logApertura2[0].LAND;
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
                PagDocum = new Vista.PagoDocumento.PagoDocumento(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                PagDocum.Show();
                this.Visibility = Visibility.Collapsed;
            }
            if (EmisionNC.IsMouseOver == true)
            {
                NotaCredit = new Vista.NotaCredito.NotaCredito(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                NotaCredit.Show();
                this.Visibility = Visibility.Collapsed;
            }

            if (Anulación.IsMouseOver == true)
            {
                Anula = new Anulacion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Anula.Show();
                this.Visibility = Visibility.Collapsed;
            }

            if (Reimpresion.IsMouseOver == true)
            {
                Reimp = new Reimpresion.Reimpresion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Reimp.Show();
                this.Visibility = Visibility.Collapsed;
            }

            if (RecaudacionVeh.IsMouseOver == true)
            {
                Vehi = new Vehiculos.Vehiculo(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Vehi.Show();
                this.Visibility = Visibility.Collapsed;
            }

            if (CierreCaja.IsMouseOver == true)
            {
                CierCaja = new CierreCaja.CierreCaja(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                CierCaja.Show();
                this.Visibility = Visibility.Collapsed;
            }
            if (ReportesCaja.IsMouseOver == true)
            {
                Reporte = new Reportes.Reportes(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Reporte.Show();
                this.Visibility = Visibility.Collapsed;
            }
        }

        private void btnBuscarAnulV_Click(object sender, RoutedEventArgs e)
        {
            txtUserAnula.Text = "";
            txtComentAnula.Text = "";
            LimpiarViasDePago();
            chkFiltro.IsChecked = false;
            btnAnularV.IsEnabled = false;
            if ((txtComprAnV.Text == "") && (txtRUTAnV.Text == ""))
            {
                System.Windows.MessageBox.Show("Ingrese un RUT o un número de comprobante");
            }
            else
            {
                ListaDocumentosAnulacionVehiculos();
            }
            GC.Collect();
        }

        public void ListaDocumentosAnulacionVehiculos()
        {
            try
            {
                BusquedaAnulacion busquedaanulacion = new BusquedaAnulacion();
                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();

                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();
                btnRevisDoc.Visibility = Visibility.Visible;

                // PartidasAbiertas partidasabiertas = new PartidasAbiertas();
                if (RBRutAnulV.IsChecked == true)
                {
                    string RUTAux = "";
                    foreach (char value in txtRUTAnV.Text.ToUpper())
                    {
                        if (value != 8207)
                        {
                            RUTAux = RUTAux + value;
                        }
                    }

                    RUTAux = RUTAux.Trim();
                    String RUT = DigitoVerificador(RUTAux);
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRUTAnV.Focus();
                    }
                    else
                    {
                        busquedaanulacion.docsanulacion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprAnV.Text, txtRUTAnV.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "V");
                    }
                }
                else if (RBDocAnulV.IsChecked == true)
                {
                    string Documento = "";
                    Documento = txtComprAnV.Text;
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
                        txtComprAnV.Text = Documento;

                        busquedaanulacion.docsanulacion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprAnV.Text, txtRUTAnV.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "V");
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }

                if (busquedaanulacion.CabeceraDocs.Count > 0)
                {
                    GBDetalleDocs.Visibility = Visibility.Visible;
                    DGDocCabec.ItemsSource = null;
                    DGDocCabec.Items.Clear();
                    List<CAB_COMPAUX> partidaopen = new List<CAB_COMPAUX>();

                    for (int k = 0; k < busquedaanulacion.CabeceraDocs.Count; k++)
                    {
                        CAB_COMPAUX partOpen = new CAB_COMPAUX();
                        partOpen.ISSELECTED = false;
                        partOpen.AUT_JEF = busquedaanulacion.CabeceraDocs[k].AUT_JEF;
                        partOpen.CLASE_DOC = busquedaanulacion.CabeceraDocs[k].CLASE_DOC;
                        partOpen.CLIENTE = busquedaanulacion.CabeceraDocs[k].CLIENTE;
                        partOpen.DESCRIPCION = busquedaanulacion.CabeceraDocs[k].DESCRIPCION;
                        partOpen.FECHA_COMP = busquedaanulacion.CabeceraDocs[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = busquedaanulacion.CabeceraDocs[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = busquedaanulacion.CabeceraDocs[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = busquedaanulacion.CabeceraDocs[k].ID_COMPROBANTE;
                        partOpen.LAND = busquedaanulacion.CabeceraDocs[k].LAND;
                        partOpen.MONEDA = busquedaanulacion.CabeceraDocs[k].MONEDA;
                        partOpen.MONTO_DOC = busquedaanulacion.CabeceraDocs[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = busquedaanulacion.CabeceraDocs[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = busquedaanulacion.CabeceraDocs[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = busquedaanulacion.CabeceraDocs[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = busquedaanulacion.CabeceraDocs[k].TXT_CLASE_DOC;
                        partOpen.VIA_PAGO = busquedaanulacion.CabeceraDocs[k].VIA_PAGO;

                        partidaopen.Add(partOpen);
                    }

                    DGDocCabec.ItemsSource = partidaopen;
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = busquedaanulacion.DetalleDocs;
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

        //FUNCION QUE LIMPIA TODOS LOS ELEMENTOS PRESENTES EN LAS VIAS DE PAGO
        private void LimpiarViasDePago()
        {
            //Limpiar aqui el resumen de las vias de pago
            GC.Collect();
        }

        private void btnBuscarAnul_Click(object sender, RoutedEventArgs e)
        {
            txtUserAnula.Text = "";
            txtComentAnula.Text = "";
            LimpiarViasDePago();
            chkFiltro.IsChecked = false;
            btnAnularV.IsEnabled = false;
            btnRevisDoc.IsEnabled = true;
            if ((txtComprAn.Text == "") && (txtRUTAn.Text == ""))
            {
                System.Windows.MessageBox.Show("Ingrese un RUT o un número de comprobante");
            }
            else
            {
                ListaDocumentosAnulacion();
            }
            GC.Collect();   
        }
        public void ListaDocumentosAnulacion()
        {
            try
            {
                BusquedaAnulacion busquedaanulacion = new BusquedaAnulacion();

                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();

                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();

                btnRevisDoc.Visibility = Visibility.Collapsed;
                if (RBRutAnul.IsChecked == true)
                {
                    string RUTAux = "";
                    foreach (char value in txtRUTAn.Text.ToUpper())
                    {
                        if (value != 8207)
                        {
                            RUTAux = RUTAux + value;
                        }
                    }

                    RUTAux = RUTAux.Trim();
                    String RUT = DigitoVerificador(RUTAux);
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRUTAn.Focus();
                    }
                    else
                    {
                        busquedaanulacion.docsanulacion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprAn.Text, txtRUTAn.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "A");
                    }
                }
                else if (RBDocAnul.IsChecked == true)
                {
                    string Documento = "";
                    Documento = txtComprAn.Text;
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
                        txtComprAn.Text = Documento;

                        busquedaanulacion.docsanulacion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprAn.Text, txtRUTAn.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), Accion);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }
                if (busquedaanulacion.CabeceraDocs.Count > 0)
                {
                    GBDetalleDocs.Visibility = Visibility.Visible;
                    DGDocCabec.ItemsSource = null;
                    DGDocCabec.Items.Clear();
                    List<CAB_COMPAUX> partidaopen = new List<CAB_COMPAUX>();

                    for (int k = 0; k < busquedaanulacion.CabeceraDocs.Count; k++)
                    {
                        CAB_COMPAUX partOpen = new CAB_COMPAUX();
                        partOpen.ISSELECTED = false;
                        partOpen.AUT_JEF = busquedaanulacion.CabeceraDocs[k].AUT_JEF;
                        partOpen.CLASE_DOC = busquedaanulacion.CabeceraDocs[k].CLASE_DOC;
                        partOpen.CLIENTE = busquedaanulacion.CabeceraDocs[k].CLIENTE;
                        partOpen.DESCRIPCION = busquedaanulacion.CabeceraDocs[k].DESCRIPCION;
                        partOpen.FECHA_COMP = busquedaanulacion.CabeceraDocs[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = busquedaanulacion.CabeceraDocs[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = busquedaanulacion.CabeceraDocs[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = busquedaanulacion.CabeceraDocs[k].ID_COMPROBANTE;
                        partOpen.LAND = busquedaanulacion.CabeceraDocs[k].LAND;
                        partOpen.MONEDA = busquedaanulacion.CabeceraDocs[k].MONEDA;
                        partOpen.MONTO_DOC = busquedaanulacion.CabeceraDocs[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = busquedaanulacion.CabeceraDocs[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = busquedaanulacion.CabeceraDocs[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = busquedaanulacion.CabeceraDocs[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = busquedaanulacion.CabeceraDocs[k].TXT_CLASE_DOC;
                        partOpen.VIA_PAGO = busquedaanulacion.CabeceraDocs[k].VIA_PAGO;

                        partidaopen.Add(partOpen);
                    }

                    if (partidaopen.Count > 0)
                    {
                        DGDocCabec.ItemsSource = partidaopen;
                        DGDocDet.ItemsSource = null;
                        DGDocDet.Items.Clear();
                        DGDocDet.ItemsSource = busquedaanulacion.DetalleDocs;
                        btnAnular.IsEnabled = true;
                        
                        DGDocCabec.Visibility = Visibility.Visible;
                        DGDocDet.Visibility = Visibility.Visible;
                        label10.Visibility = Visibility.Collapsed;
                        btnRevisDoc.Visibility = Visibility.Visible;
                        btnAnular.IsEnabled = false;

                        if (partidaopen[0].VIA_PAGO == "E" || partidaopen[0].VIA_PAGO == "R" || partidaopen[0].VIA_PAGO == "N" || partidaopen[0].VIA_PAGO == "U")
                        {
                            btnRevisDoc.IsEnabled = true;
                        }
                        else
                        {
                            btnRevisDoc.IsEnabled = false;
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
        private void tabControlAnulacion_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {          
            txtComprAnV.Text = "";
            txtComprAn.Text = "";
            txtRUTAnV.Text = "";
            txtRUTAn.Text = "";
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            GC.Collect();
        }

        private void RBRutAnul_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTAn.Text = "";
            txtRUTAn.Visibility = Visibility.Visible;
            txtComprAn.Text = "";
            txtComprAn.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            detalle = new List<DET_COMP>();
            cabecera = new List<CAB_COMP>();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnBuscarAnul.IsEnabled = true;
            btnAnular.IsEnabled = false;
            GC.Collect();
        }

        private void RBDocAnul_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTAn.Text = "";
            txtRUTAn.Visibility = Visibility.Collapsed;
            txtComprAn.Text = "";
            txtComprAn.Visibility = Visibility.Visible;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            detalle = new List<DET_COMP>();
            cabecera = new List<CAB_COMP>();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnBuscarAnul.IsEnabled = true;
            btnAnular.IsEnabled = false;
            GC.Collect();
        }

        //ANULACION DE COMPROBANTES
        private void btnAnular_Click(object sender, RoutedEventArgs e)
        {
            if (RbAnularRecau.IsChecked == true)
            {
                Accion = "A";

                List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CAB_COMP> partidaopen = new List<CAB_COMP>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CAB_COMP partOpen = new CAB_COMP();
                        partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                        IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                        partidaopen.Add(partOpen);
                    }
                }
                if (partidaopen.Count == 0)
                {
                    System.Windows.MessageBox.Show("Seleccione un comprobante en la tabla de cabeceras");
                }
                else
                {
                     
                    //if(montoEfecCaja > montoDocu ) { 

                    //RFC PARA ANULAR COMPROBANTES
                    AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                    anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), "A");
                    //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                    if (anulacioncomprobantes.Mensaje != "")
                    {
                        System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                    }
                    if (anulacioncomprobantes.errormessage != "")
                    {
                        System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                    }
                    if (anulacioncomprobantes.Retorno.Count > 0)
                    {
                        DGDocCabec.ItemsSource = null;
                        DGDocCabec.Items.Clear();
                        DGDocDet.ItemsSource = null;
                        DGDocDet.Items.Clear();

                        txtUserAnula.Text = "";
                        txtComentAnula.Text = "";
                        txtComprAn.Text = "";
                        txtRUTAn.Text = "";
                        btnAnular.IsEnabled = false;

                        ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                    }
                    //}else
                    //{
                    //    System.Windows.MessageBox.Show("No tiene efectivo disponible en caja, favor dirigirse al Departamento de Cobranzas.", "Mensaje");
                    //}
                }
            }
            else if (RbAnularNegocio.IsChecked == true)
            {
                Accion = "D";

                List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CAB_COMP> partidaopen = new List<CAB_COMP>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CAB_COMP partOpen = new CAB_COMP();
                        partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                        partOpen.VIA_PAGO = DocsAPagar[k].VIA_PAGO;
                        IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                        partidaopen.Add(partOpen);
                    }
                }

                montoDocu = Convert.ToDouble(partidaopen[0].MONTO_DOC.Replace(".",""));

                if (txtTotEfect.Text == "")
                {
                    montoEfecCaja = 0;
                }
                else
                {
                    montoEfecCaja = Convert.ToDouble(txtTotEfect.Text.Replace(".",""));
                }

                if (partidaopen[0].VIA_PAGO == "N" || partidaopen[0].VIA_PAGO == "R" || partidaopen[0].VIA_PAGO == "U" || partidaopen[0].VIA_PAGO == "E")
                {
                    if (montoEfecCaja >= montoDocu)
                    {
                        //RFC PARA ANULAR COMPROBANTES
                        AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                        anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                        //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                        if (anulacioncomprobantes.Mensaje != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                        }
                        if (anulacioncomprobantes.errormessage != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                        }
                        if (anulacioncomprobantes.Retorno.Count > 0)
                        {
                            DGDocCabec.ItemsSource = null;
                            DGDocCabec.Items.Clear();
                            DGDocDet.ItemsSource = null;
                            DGDocDet.Items.Clear();

                            txtUserAnula.Text = "";
                            txtComentAnula.Text = "";
                            txtComprAn.Text = "";
                            txtRUTAn.Text = "";
                            btnAnular.IsEnabled = false;

                            ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                        }
                    }
                    else
                    {
                        //  System.Windows.MessageBox.Show("No tiene efectivo disponible en caja, favor dirigirse al Departamento de Cobranzas.", "Mensaje");
                        //RFC PARA ANULAR COMPROBANTES
                        AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                        anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                        //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                        if (anulacioncomprobantes.Mensaje != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                        }
                        if (anulacioncomprobantes.errormessage != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                        }
                        if (anulacioncomprobantes.Retorno.Count > 0)
                        {
                            DGDocCabec.ItemsSource = null;
                            DGDocCabec.Items.Clear();
                            DGDocDet.ItemsSource = null;
                            DGDocDet.Items.Clear();

                            txtUserAnula.Text = "";
                            txtComentAnula.Text = "";
                            txtComprAn.Text = "";
                            txtRUTAn.Text = "";
                            btnAnular.IsEnabled = false;

                            ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                        }
                    }   
                 }
                else
                {
                   
                        //RFC PARA ANULAR COMPROBANTES
                        AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                        anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                        //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                        if (anulacioncomprobantes.Mensaje != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                        }
                        if (anulacioncomprobantes.errormessage != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                        }
                        if (anulacioncomprobantes.Retorno.Count > 0)
                        {
                            DGDocCabec.ItemsSource = null;
                            DGDocCabec.Items.Clear();
                            DGDocDet.ItemsSource = null;
                            DGDocDet.Items.Clear();

                            txtUserAnula.Text = "";
                            txtComentAnula.Text = "";
                            txtComprAn.Text = "";
                            txtRUTAn.Text = "";
                            btnAnular.IsEnabled = false;

                            ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                        }
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Debe Seleccionar un Tipo Anulacion", "Mensaje");
            }
            GC.Collect();
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
                        , Convert.ToString(textBlock7.Content), Caja, Referencia, Documento, DocContable, InOut, "", Pedido, txtMandante.Text);
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
        private void btnAnularV_Click(object sender, RoutedEventArgs e)
        {
            if (RbAnularNegocioVehi.IsChecked == true)
            {
                Accion = "D";
                //BUSQUEDA DEL COMPROBANTE SELECCIONADO
                string IdComprobante = "";
                List<CAB_COMP> Comprobante = new List<CAB_COMP>();
                List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CAB_COMP> partidaopen = new List<CAB_COMP>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CAB_COMP partOpen = new CAB_COMP();
                        partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                        IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.VIA_PAGO = DocsAPagar[k].VIA_PAGO;
                        partidaopen.Add(partOpen);
                    }
                }
                if (partidaopen.Count == 0)
                {
                    System.Windows.MessageBox.Show("Seleccione un comprobante en la tabla de cabeceras");
                }
                else
                {

                    //montoDocu = Convert.ToDouble(partidaopen[0].MONTO_DOC.Replace(".", ""));

                    //if (txtTotEfect2.Text == "")
                    //{
                    //    montoEfecCaja = 0;
                    //}
                    //else
                    //{
                    //    montoEfecCaja = Convert.ToDouble(txtTotEfect2.Text.Replace(".", ""));
                    //}

                    //if (partidaopen[0].VIA_PAGO == "N" || partidaopen[0].VIA_PAGO == "R" || partidaopen[0].VIA_PAGO == "U" || partidaopen[0].VIA_PAGO == "E")
                    //{
                    //if (montoEfecCaja > montoDocu)
                    //{
                    //RFC PARA ANULAR COMPROBANTES

                    MessageBoxResult result = System.Windows.MessageBox.Show("¿Con esta acción anulara todas las Centralizaciones esta seguro?", "Mensaje", MessageBoxButton.YesNo);
                    switch (result)
                    {
                        case MessageBoxResult.Yes:

                            AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                            anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, txtUserAnula.Text, txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                            //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                            if (anulacioncomprobantes.Mensaje != "")
                            {
                                System.Windows.Forms.MessageBox.Show(anulacioncomprobantes.Mensaje);
                                lbvehi.Visibility = Visibility.Collapsed;
                                txtTotEfect2.Visibility = Visibility.Collapsed;
                            }
                            if (anulacioncomprobantes.errormessage != "")
                            {
                                System.Windows.Forms.MessageBox.Show(anulacioncomprobantes.errormessage);
                                lbvehi.Visibility = Visibility.Collapsed;
                                txtTotEfect2.Visibility = Visibility.Collapsed;
                            }

                            if (anulacioncomprobantes.Retorno.Count > 0)
                            {
                                ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                                DGDocCabec.ItemsSource = null;
                                DGDocCabec.Items.Clear();
                                DGDocDet.ItemsSource = null;
                                DGDocDet.Items.Clear();
                                txtUserAnulaV.Text = "";
                                txtComentAnulaV.Text = "";
                                txtComprAnV.Text = "";
                                txtRUTAnV.Text = "";
                                btnAnularV.IsEnabled = false;
                            }
                            break;

                        case MessageBoxResult.No:
                            DGDocCabec.ItemsSource = null;
                            DGDocCabec.Items.Clear();
                            DGDocDet.ItemsSource = null;
                            DGDocDet.Items.Clear();
                            txtUserAnulaV.Text = "";
                            txtComentAnulaV.Text = "";
                            txtComprAnV.Text = "";
                            txtRUTAnV.Text = "";
                            btnAnularV.IsEnabled = false;
                            lbvehi.Visibility = Visibility.Collapsed;
                            txtTotEfect2.Visibility = Visibility.Collapsed; 
                            break;
                    }                          
                        //}
                        //else
                        //{
                        //    System.Windows.MessageBox.Show("No tiene efectivo disponible en caja, favor dirigirse al Departamento de Cobranzas.", "Mensaje");
                        //}
                    }
                    //else
                    //{
                    //    //RFC PARA ANULAR COMPROBANTES
                    //    AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                    //    anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, txtUserAnula.Text, txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                    //    //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                    //    if (anulacioncomprobantes.Mensaje != "")
                    //    {
                    //        System.Windows.Forms.MessageBox.Show(anulacioncomprobantes.Mensaje);
                    //    }
                    //    if (anulacioncomprobantes.errormessage != "")
                    //    {
                    //        System.Windows.Forms.MessageBox.Show(anulacioncomprobantes.errormessage);
                    //    }

                    //    if (anulacioncomprobantes.Retorno.Count > 0)
                    //    {
                    //        ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                    //        DGDocCabec.ItemsSource = null;
                    //        DGDocCabec.Items.Clear();
                    //        DGDocDet.ItemsSource = null;
                    //        DGDocDet.Items.Clear();
                    //        txtUserAnulaV.Text = "";
                    //        txtComentAnulaV.Text = "";
                    //        txtComprAnV.Text = "";
                    //        txtRUTAnV.Text = "";
                    //        btnAnularV.IsEnabled = false;
                    //    }
                    //}
            }
            else if (RbAnularNegocio2.IsChecked == true)
            {
                Accion = "A";

                List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CAB_COMP> partidaopen = new List<CAB_COMP>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CAB_COMP partOpen = new CAB_COMP();
                        partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                        partOpen.VIA_PAGO = DocsAPagar[k].VIA_PAGO;
                        IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                        partidaopen.Add(partOpen);
                    }
                }

                //montoDocu = Convert.ToDouble(partidaopen[0].MONTO_DOC.Replace(".", ""));

                //if (txtTotEfect.Text == "")
                //{
                //    montoEfecCaja = 0;
                //}
                //else
                //{
                //    montoEfecCaja = Convert.ToDouble(txtTotEfect.Text.Replace(".", ""));
                //}

                //if (partidaopen[0].VIA_PAGO == "N" || partidaopen[0].VIA_PAGO == "R" || partidaopen[0].VIA_PAGO == "U" || partidaopen[0].VIA_PAGO == "E")
                //{
                    //if (montoEfecCaja >= montoDocu)
                    //{
                        //RFC PARA ANULAR COMPROBANTES
                        AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                        anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                        //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                        if (anulacioncomprobantes.Mensaje != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                            lbvehi.Visibility = Visibility.Collapsed;
                            txtTotEfect2.Visibility = Visibility.Collapsed;
                         }
                        if (anulacioncomprobantes.errormessage != "")
                        {
                            System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                            lbvehi.Visibility = Visibility.Collapsed;
                            txtTotEfect2.Visibility = Visibility.Collapsed;
                         }
                        if (anulacioncomprobantes.Retorno.Count > 0)
                        {
                            DGDocCabec.ItemsSource = null;
                            DGDocCabec.Items.Clear();
                            DGDocDet.ItemsSource = null;
                            DGDocDet.Items.Clear();
                            txtUserAnula.Text = "";
                            txtComentAnula.Text = "";
                            txtComprAn.Text = "";
                            txtRUTAn.Text = "";
                            btnAnular.IsEnabled = false;

                            ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                        }
                    //}
                    //else
                    //{
                    //    //  System.Windows.MessageBox.Show("No tiene efectivo disponible en caja, favor dirigirse al Departamento de Cobranzas.", "Mensaje");
                    //    //RFC PARA ANULAR COMPROBANTES
                    //    AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                    //    anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                    //    //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                    //    if (anulacioncomprobantes.Mensaje != "")
                    //    {
                    //        System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                    //    }
                    //    if (anulacioncomprobantes.errormessage != "")
                    //    {
                    //        System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                    //    }
                    //    if (anulacioncomprobantes.Retorno.Count > 0)
                    //    {
                    //        DGDocCabec.ItemsSource = null;
                    //        DGDocCabec.Items.Clear();
                    //        DGDocDet.ItemsSource = null;
                    //        DGDocDet.Items.Clear();

                    //        txtUserAnula.Text = "";
                    //        txtComentAnula.Text = "";
                    //        txtComprAn.Text = "";
                    //        txtRUTAn.Text = "";
                    //        btnAnular.IsEnabled = false;

                    //        ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                    //    }
                    //}
                //}
                //else
                //{

                //    //RFC PARA ANULAR COMPROBANTES
                //    AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                //    anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, Convert.ToString(textBlock7.Content), txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Accion);
                //    //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                //    if (anulacioncomprobantes.Mensaje != "")
                //    {
                //        System.Windows.MessageBox.Show(anulacioncomprobantes.Mensaje);
                //    }
                //    if (anulacioncomprobantes.errormessage != "")
                //    {
                //        System.Windows.MessageBox.Show(anulacioncomprobantes.errormessage);
                //    }
                //    if (anulacioncomprobantes.Retorno.Count > 0)
                //    {
                //        DGDocCabec.ItemsSource = null;
                //        DGDocCabec.Items.Clear();
                //        DGDocDet.ItemsSource = null;
                //        DGDocDet.Items.Clear();

                //        txtUserAnula.Text = "";
                //        txtComentAnula.Text = "";
                //        txtComprAn.Text = "";
                //        txtRUTAn.Text = "";
                //        btnAnular.IsEnabled = false;

                //        ImpresionesDeDocumentosAutomaticas(anulacioncomprobantes.NumComprobante, "X");
                //    }
                //}
            }
            else
            {
                System.Windows.MessageBox.Show("Debe Seleccionar un Tipo Anulacion", "Mensaje");
            }

            GC.Collect();
        }

        private void btnAutAnul_Click(object sender, RoutedEventArgs e)
        {
            List<CAB_COMP> Comprobante = new List<CAB_COMP>();
            for (int i = 0; i < DGDocCabec.SelectedItems.Count; i++)
            {
                {
                    Comprobante.Add(DGDocCabec.SelectedItems[i] as CAB_COMP);
                }
            }
            if (Comprobante.Count > 0)
            {
                CheckUserAnulacion marcardocsparaanular = new CheckUserAnulacion();
                marcardocsparaanular.checkdocsanulacion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text
                , txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content)
                , Convert.ToString(lblPais.Content), txtRUTAnV.Text, txtComprAnV.Text, Convert.ToString(lblSociedad.Content), "A", Comprobante);

                if (marcardocsparaanular.errormessage != "")
                {
                    System.Windows.Forms.MessageBox.Show(marcardocsparaanular.errormessage);
                }
                if (marcardocsparaanular.message != "")
                {
                    System.Windows.Forms.MessageBox.Show(marcardocsparaanular.message);
                }
                if (marcardocsparaanular.valido == "X")
                {
                    btnAnularV.IsEnabled = true;
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Seleccione la vía de pago a depositar");
            }
        }
        private void btnRevisDoc_Click(object sender, RoutedEventArgs e)
        {
           
            List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
            DocsAPagar.Clear();
            if (this.DGDocCabec.Items.Count > 0)
            {
                for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                {
                    if (i == 0)
                        DGDocCabec.Items.MoveCurrentToFirst();
                    {
                        DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                    }
                    DGDocCabec.Items.MoveCurrentToNext();
                }
            }

            List<CAB_COMP> partidaopen = new List<CAB_COMP>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CAB_COMP partOpen = new CAB_COMP();
                        partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                        partOpen.VIA_PAGO = DocsAPagar[k].VIA_PAGO;
                         IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                        partidaopen.Add(partOpen);

                    }
                }

            AnulacionComprobantes anula = new AnulacionComprobantes();

             if (partidaopen[0].VIA_PAGO == "N" || partidaopen[0].VIA_PAGO == "R" || partidaopen[0].VIA_PAGO == "U" || partidaopen[0].VIA_PAGO == "E")
            {
                ViaPago = "E";
                
                anula.MontoEfectivo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content), logApertura2[0].ID_APERTURA, SociedCaja.Text, ViaPago);
            }
            //else
            //{
            //    ViaPago = partidaopen[0].VIA_PAGO;

            //    anula.MontoEfectivo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content), logApertura2[0].ID_APERTURA, SociedCaja.Text, ViaPago);
            //}
             
            if(anula.Efectivo != "00")
            {
                if (tabItem5.IsSelected == false)
                {
                    label41.Visibility = Visibility.Visible;
                    txtTotEfect.Visibility = Visibility.Visible;
                    txtTotEfect.Text = anula.Efectivo;
                    btnAnular.IsEnabled = true;
                }
                else
                {
                    txtTotEfect2.Text = anula.Efectivo;
                    lbvehi.Visibility = Visibility.Visible;
                    txtTotEfect2.Visibility = Visibility;
                    btnAnularV.IsEnabled = true;
                }
                
                label10.Visibility = Visibility.Visible;           
            }
            else
            {
                txtTotEfect.Text = "0";
                if (tabItem5.IsSelected == false)
                {
                    label41.Visibility = Visibility.Collapsed;
                    txtTotEfect.Visibility = Visibility.Collapsed;
                    btnAnular.IsEnabled = false;
                }
                else
                {
                    lbvehi.Visibility = Visibility.Collapsed;
                    txtTotEfect2.Visibility = Visibility.Collapsed;
                    btnAnularV.IsEnabled = false;
                }
                label10.Visibility = Visibility.Visible;               
                System.Windows.Forms.MessageBox.Show("No tiene efectivo disponible en caja, favor dirigirse al Departamento de Cobranzas.", "Mensaje");
            }

            GC.Collect();
        }

        private void ChkDetalleDocs_Checked(object sender, RoutedEventArgs e)
        {
            if (GBAnulacion.IsVisible == true)
            {
                List<CAB_COMPAUX> partidaseleccionadasaux2 = new List<CAB_COMPAUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                {
                    partidaseleccionadasaux2[k].ISSELECTED = true;
                }
                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();
                DGDocCabec.ItemsSource = partidaseleccionadasaux2;
            }
            GC.Collect();
        }

        private void ChkDetalleDocs_Unchecked(object sender, RoutedEventArgs e)
        {
            if (GBAnulacion.IsVisible == true)
            {
                List<CAB_COMPAUX> partidaseleccionadasaux2 = new List<CAB_COMPAUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                {
                    partidaseleccionadasaux2[k].ISSELECTED = true;
                }
                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();
                DGDocCabec.ItemsSource = partidaseleccionadasaux2;
            }           
        }

        private void chkFiltro_Checked(object sender, RoutedEventArgs e)
        {
            if (GBAnulacion.IsVisible)
            {
                detalleaux.Clear();
                List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CAB_COMP> partidaopen = new List<CAB_COMP>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CAB_COMP partOpen = new CAB_COMP();
                        partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                        partidaopen.Add(partOpen);
                    }
                }
                if (partidaopen.Count > 0)
                {
                    List<DET_COMP> DetalleDocs = new List<DET_COMP>();
                    DetalleDocs.Clear();
                    for (int i = 1; i <= DGDocDet.Items.Count; i++)
                    {
                        if (i == 1)
                        {
                            DGDocDet.Items.MoveCurrentToFirst();
                        }
                        if (DGDocDet.Items.CurrentItem != null)
                        {
                            DetalleDocs.Add(DGDocDet.Items.CurrentItem as DET_COMP);
                        }

                        DGDocDet.Items.MoveCurrentToNext();
                    }
                    detalle.Clear();
                    detalleaux.Clear();
                    for (int i = 0; i < DetalleDocs.Count; i++)
                    {
                        detalle.Add(DetalleDocs[i]);
                        if (partidaopen[0].ID_COMPROBANTE == DetalleDocs[i].ID_COMPROBANTE)
                        {
                            detalleaux.Add(DetalleDocs[i]);
                        }
                    }
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = detalleaux;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Seleccione un registro en la tabla de documentos/cabecera");
                }
            }
            GC.Collect();
        }

        private void chkFiltro_UnChecked(object sender, RoutedEventArgs e)
        {
            if (GBAnulacion.IsVisible)
            {
                if (detalle.Count > 0)
                {
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = detalle;
                    detalleaux.Clear();
                }
            }
            GC.Collect();
        }
        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {
            GC.Collect();
        }

        private void DGMonitor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            timer.Stop();
            GC.Collect();
        }

        private void RBDocAnulV_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTAnV.Text = "";
            txtRUTAnV.Visibility = Visibility.Collapsed;
            txtComprAnV.Text = "";
            txtComprAnV.Visibility = Visibility.Visible;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            detalle = new List<DET_COMP>();
            cabecera = new List<CAB_COMP>();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnBuscarAnulV.IsEnabled = true;
            GC.Collect();
        }
        private void RBRutAnulV_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTAnV.Text = "";
            txtRUTAnV.Visibility = Visibility.Visible;
            txtComprAnV.Text = "";
            txtComprAnV.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            detalle = new List<DET_COMP>();
            cabecera = new List<CAB_COMP>();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnBuscarAnulV.IsEnabled = true;
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

        private void Anulación_Click(object sender, RoutedEventArgs e)
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

        private void CierreCaja_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void ReportesCaja_Click(object sender, RoutedEventArgs e)
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

        private void RbAnularNegocio_Checked(object sender, RoutedEventArgs e)
        {
            List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
            DocsAPagar.Clear();
            if (this.DGDocCabec.Items.Count > 0)
            {
                for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                {
                    if (i == 0)
                        DGDocCabec.Items.MoveCurrentToFirst();
                    {
                        DocsAPagar.Add(DGDocCabec.Items.CurrentItem as CAB_COMPAUX);
                    }
                    DGDocCabec.Items.MoveCurrentToNext();
                }
            }

            List<CAB_COMP> partidaopen = new List<CAB_COMP>();

            for (int k = 0; k < DocsAPagar.Count; k++)
            {
                if (DocsAPagar[k].ISSELECTED == true)
                {
                    CAB_COMP partOpen = new CAB_COMP();
                    partOpen.AUT_JEF = DocsAPagar[k].AUT_JEF;
                    partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                    partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                    partOpen.DESCRIPCION = DocsAPagar[k].DESCRIPCION;
                    partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                    partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                    partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                    partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                    partOpen.LAND = DocsAPagar[k].LAND;
                    partOpen.MONEDA = DocsAPagar[k].MONEDA;
                    partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                    partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                    partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                    partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                    partOpen.TXT_CLASE_DOC = DocsAPagar[k].TXT_CLASE_DOC;
                    partOpen.VIA_PAGO = DocsAPagar[k].VIA_PAGO;
                    IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                    partidaopen.Add(partOpen);

                    //ViaPago = "E";
                    if (partidaopen[k].VIA_PAGO == "N" || partidaopen[k].VIA_PAGO == "R" || partidaopen[k].VIA_PAGO == "U" || partidaopen[k].VIA_PAGO == "E")
                    {
                        btnRevisDoc.IsEnabled = true;
                    }
                    else
                    {
                        btnRevisDoc.IsEnabled = false;
                        btnAnular.IsEnabled = true;
                    }
                }
            }
         
            GC.Collect();
            
            label41.Visibility = Visibility.Collapsed;
            txtTotEfect.Visibility = Visibility.Collapsed;
            txtTotEfect.Text = "";
        }

        private void RbAnularRecau_Checked(object sender, RoutedEventArgs e)
        {
            btnRevisDoc.IsEnabled = false;
            btnAnular.IsEnabled = true;
            label41.Visibility = Visibility.Collapsed;
            txtTotEfect.Visibility = Visibility.Collapsed;
            txtTotEfect.Text = "";
        }

        private void RbAnularNegocio2_Checked(object sender, RoutedEventArgs e)
        {

        }
    }
}
