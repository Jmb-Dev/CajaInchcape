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


namespace CajaIndigo.Vista.Reimpresion
{
    public partial class Reimpresion : System.Windows.Window
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


        Vista.NotaCredito.NotaCredito NotaCredit;
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.PagoDocumento.PagoDocumento PagDocum;
        Vista.Reportes.Reportes Reporte;

        List<LOG_APERTURA> logApertura = new List<LOG_APERTURA>();
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();

        public Reimpresion()
        {
            InitializeComponent();
        }

        public Reimpresion(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
        {
            try
            {
                InitializeComponent();
                int test = 0;
              //  GBInicio.Visibility = Visibility.Visible;
                GBMonitor.Visibility = Visibility.Visible;
                //GBInicio.Visibility = Visibility.Collapsed;
                GBReimpresion.Visibility = Visibility.Visible;
                GBDetalleDocs.Visibility = Visibility.Collapsed;
                GBCommentCierre.Visibility = Visibility.Collapsed;
                textBlock6.Content = cajaconect;
                textBlock7.Content = usuariologg;
                textBlock8.Content = sucursal;
                textBlock9.Content = usuariotemp ;
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

        //FUNCION QUE TRAE LOS DOCUMENTOS A REIMPRIMIR
        private void ListaDocumentosReimpresion()
        {
            try
            {
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }

                BusquedaReimpresiones busquedareimpresiones = new BusquedaReimpresiones();
                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();
                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();
                btnRevisDoc.Visibility = Visibility.Collapsed;
                if (RBRutReimp.IsChecked == true)
                {
                    string RUTAux = "";
                    foreach (char value in txtRUTReimp.Text.ToUpper())
                    {
                        if (value != 8207)
                        {
                            RUTAux = RUTAux + value;
                        }
                    }
                    RUTAux = RUTAux.Trim();
                    //***RFC Partidas abiertas para REIMPRESION busqueda por RUT
                    String RUT = DigitoVerificador(RUTAux);
                    //Verificacion de RUT
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRUTReimp.Focus();
                    }
                    else
                    {
                        busquedareimpresiones.docsreimpresion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprReimp.Text, txtRUTReimp.Text, logApertura2[0].ID_REGISTRO, Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "", Convert.ToString(lblSociedad.Content));
                    }
                }
                else if (RBDocReimp.IsChecked == true)
                {
                    //***RFC Partidas abiertas para REIMPRESION busqueda por numero de documento
                    string Documento = "";
                    Documento = txtComprReimp.Text;
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
                        txtComprReimp.Text = Documento;

                        busquedareimpresiones.docsreimpresion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprReimp.Text, txtRUTReimp.Text, logApertura2[0].ID_REGISTRO, Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "", Convert.ToString(lblSociedad.Content));
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }
               
                if (busquedareimpresiones.Documentos.Count > 0)
                {
                    GBDetalleDocs.Visibility = Visibility.Visible;
                    DGDocCabec.ItemsSource = null;
                    DGDocCabec.Items.Clear();
                    List<DOCUMENTOSAUX> partidaopen = new List<DOCUMENTOSAUX>();

                    for (int k = 0; k < busquedareimpresiones.Documentos.Count; k++)
                    {
                        DOCUMENTOSAUX partOpen = new DOCUMENTOSAUX();
                        partOpen.ISSELECTED = false;
                        partOpen.ACC = busquedareimpresiones.Documentos[k].ACC;
                        partOpen.CEBE = busquedareimpresiones.Documentos[k].CEBE;
                        partOpen.CLASE_CUENTA = busquedareimpresiones.Documentos[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = busquedareimpresiones.Documentos[k].CLASE_DOC;
                        partOpen.CME = busquedareimpresiones.Documentos[k].CME;
                        partOpen.APROBADOR_ANULA = busquedareimpresiones.Documentos[k].APROBADOR_ANULA;
                        partOpen.APROBADOR_EX = busquedareimpresiones.Documentos[k].APROBADOR_EX;
                        partOpen.CAJERO_GEN = busquedareimpresiones.Documentos[k].CAJERO_GEN;
                        partOpen.CAJERO_RESP = busquedareimpresiones.Documentos[k].CAJERO_RESP;
                        partOpen.CLIENTE = busquedareimpresiones.Documentos[k].CLIENTE;
                        partOpen.FECHA_DOC = busquedareimpresiones.Documentos[k].FECHA_DOC;
                        partOpen.EXCEPCION = busquedareimpresiones.Documentos[k].EXCEPCION;
                        partOpen.FECHA_COMP = busquedareimpresiones.Documentos[k].FECHA_COMP;
                        partOpen.MONEDA = busquedareimpresiones.Documentos[k].MONEDA;
                        partOpen.FECHA_DOC = busquedareimpresiones.Documentos[k].FECHA_DOC;
                        partOpen.FECHA_VENC_DOC = busquedareimpresiones.Documentos[k].FECHA_VENC_DOC;
                        partOpen.HORA = busquedareimpresiones.Documentos[k].HORA;
                        partOpen.MONTO_DIFERENCIA = busquedareimpresiones.Documentos[k].MONTO_DIFERENCIA;
                        partOpen.MONTO_DOC = busquedareimpresiones.Documentos[k].MONTO_DOC;
                        partOpen.NOTA_VENTA = busquedareimpresiones.Documentos[k].NOTA_VENTA;
                        partOpen.NRO_ANULACION = busquedareimpresiones.Documentos[k].NRO_ANULACION;
                        partOpen.NRO_COMPENSACION = busquedareimpresiones.Documentos[k].NRO_COMPENSACION;
                        partOpen.NRO_DOCUMENTO = busquedareimpresiones.Documentos[k].NRO_DOCUMENTO;
                        partOpen.NRO_REFERENCIA = busquedareimpresiones.Documentos[k].NRO_REFERENCIA;
                        partOpen.SOCIEDAD = busquedareimpresiones.Documentos[k].SOCIEDAD;
                        partOpen.NULO = busquedareimpresiones.Documentos[k].NULO;
                        partOpen.NUM_CANCELACION = busquedareimpresiones.Documentos[k].NUM_CANCELACION;
                        partOpen.NUM_CUOTA = busquedareimpresiones.Documentos[k].NUM_CUOTA;
                        partOpen.ID_CAJA = busquedareimpresiones.Documentos[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = busquedareimpresiones.Documentos[k].ID_COMPROBANTE;
                        partOpen.LAND = busquedareimpresiones.Documentos[k].LAND;
                        partOpen.PARCIAL = busquedareimpresiones.Documentos[k].PARCIAL;
                        partOpen.POSICION = busquedareimpresiones.Documentos[k].POSICION;
                        partOpen.TEXTO_CABECERA = busquedareimpresiones.Documentos[k].TEXTO_CABECERA;
                        partOpen.TEXTO_EXCEPCION = busquedareimpresiones.Documentos[k].TEXTO_EXCEPCION;
                        partOpen.TIME = busquedareimpresiones.Documentos[k].TIME;
                        partOpen.TXT_ANULACION = busquedareimpresiones.Documentos[k].TXT_ANULACION;
                        partOpen.USR_ANULADOR = busquedareimpresiones.Documentos[k].USR_ANULADOR;
                        partidaopen.Add(partOpen);
                    }

                    DGDocCabec.ItemsSource = partidaopen;
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = busquedareimpresiones.ViasPago;
                    btnReimpr.IsEnabled = true;
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

        //BUSQUEDA DE RUT DE DOCUMENTO A REIMPRIMIR
        private void btnBuscarReimp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                chkFiltro.IsChecked = false;
                if ((txtRUTReimp.Text == "") && (txtComprReimp.Text == ""))
                {
                    System.Windows.MessageBox.Show("Ingrese un RUT o un número de comprobante");
                }
                else
                {
                    if (chkDocFiscales.IsChecked == true)
                    {
                        GBDetalleDocs.Visibility = Visibility.Visible;

                        ReimpresionFiscal reimpresionfiscal = new ReimpresionFiscal();
                        reimpresionfiscal.ReipresionFiscal2(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprReimp.Text, Convert.ToString(lblSociedad.Content));
                        objImpr = reimpresionfiscal.reimprFiscal2;

                        List<DTE_SII> url = new List<DTE_SII>();
                        for (int i = 0; i < objImpr.Count; i++)
                        {
                            DTE_SII partOpen = new DTE_SII();
                            partOpen.ISSELECTED = false;
                            partOpen.VBELN = objImpr[i].VBELN;
                            partOpen.KONDA = objImpr[i].KONDA;
                            partOpen.BUKRS = objImpr[i].BUKRS;
                            partOpen.XBLNR = objImpr[i].XBLNR;
                            partOpen.ZUONR = objImpr[i].ZUONR;
                            partOpen.TDSII = objImpr[i].TDSII;
                            partOpen.FODOC = objImpr[i].FODOC;
                            partOpen.WAERS = objImpr[i].WAERS;
                            partOpen.FECIMP = objImpr[i].FECIMP;
                            partOpen.HORIM = objImpr[i].HORIM;
                            partOpen.URLSII = objImpr[i].URLSII;
                            url.Add(partOpen);
                        }

                        DGDocCabec.ItemsSource = url;
                        btnReimpr.Visibility = Visibility.Collapsed;
                        btnReimpr.IsEnabled = false;
                        DGDocDet.Visibility = Visibility.Collapsed;
                        btnReimpr2.Visibility = Visibility.Visible;
                        btnReimpr2.IsEnabled = true;
                        btnRevisDoc.Visibility = Visibility.Collapsed;
                        chkFiltro.Visibility = Visibility.Collapsed;
                        label10.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        ListaDocumentosReimpresion();
                        DGDocDet.Visibility = Visibility.Visible;
                        btnReimpr.Visibility = Visibility.Visible;
                        btnReimpr.IsEnabled = true;
                        btnReimpr2.Visibility = Visibility.Collapsed;
                        btnReimpr2.IsEnabled = false;
                    }
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
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

        //BOTON QUE PERMITE HACER REIMPRESIONES DE COMPROBANTES
        public void btnReimpr_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                chkFiltro_Checked(chkFiltro, e);
                //LISTA DE DATOS DE VIAS DE PAGOS DEL DATAGRID DGDOCDET
                List<VIAS_PAGO2> ListViasPagos = new List<VIAS_PAGO2>();
                ListViasPagos.Clear();
                for (int i = 1; i <= DGDocDet.Items.Count; i++)
                {
                    if (i == 1)
                    {
                        DGDocDet.Items.MoveCurrentToFirst();
                    }
                    if (DGDocDet.Items.CurrentItem != null)
                    {
                        ListViasPagos.Add(DGDocDet.Items.CurrentItem as VIAS_PAGO2);
                    }

                    DGDocDet.Items.MoveCurrentToNext();
                }
                //LISTA DE DATOS DE LOS DOCUMENTOS A PAGAR DEL DATAGRID DGDOCCABEC
                List<DOCUMENTOSAUX> DocsAPagar = new List<DOCUMENTOSAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as DOCUMENTOSAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<DOCUMENTOSAUX> partidaopen = new List<DOCUMENTOSAUX>();
                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        DOCUMENTOSAUX partOpen = new DOCUMENTOSAUX();
                        partOpen.ACC = DocsAPagar[k].ACC;
                        partOpen.CEBE = DocsAPagar[k].CEBE;
                        partOpen.CLASE_CUENTA = DocsAPagar[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CME = DocsAPagar[k].CME;
                        partOpen.APROBADOR_ANULA = DocsAPagar[k].APROBADOR_ANULA;
                        partOpen.APROBADOR_EX = DocsAPagar[k].APROBADOR_EX;
                        partOpen.CAJERO_GEN = DocsAPagar[k].CAJERO_GEN;
                        partOpen.CAJERO_RESP = DocsAPagar[k].CAJERO_RESP;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                        partOpen.EXCEPCION = DocsAPagar[k].EXCEPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.HORA = DocsAPagar[k].HORA;
                        partOpen.MONTO_DIFERENCIA = DocsAPagar[k].MONTO_DIFERENCIA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NOTA_VENTA = DocsAPagar[k].NOTA_VENTA;
                        partOpen.NRO_ANULACION = DocsAPagar[k].NRO_ANULACION;
                        partOpen.NRO_COMPENSACION = DocsAPagar[k].NRO_COMPENSACION;
                        partOpen.NRO_DOCUMENTO = DocsAPagar[k].NRO_DOCUMENTO;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.SOCIEDAD = DocsAPagar[k].SOCIEDAD;
                        partOpen.NULO = DocsAPagar[k].NULO;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.NUM_CUOTA = DocsAPagar[k].NUM_CUOTA;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.PARCIAL = DocsAPagar[k].PARCIAL;
                        partOpen.POSICION = DocsAPagar[k].POSICION;
                        partOpen.TEXTO_CABECERA = DocsAPagar[k].TEXTO_CABECERA;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TIME = DocsAPagar[k].TIME;
                        partOpen.TXT_ANULACION = DocsAPagar[k].TXT_ANULACION;
                        partOpen.USR_ANULADOR = DocsAPagar[k].USR_ANULADOR;
                        partidaopen.Add(partOpen);
                    }
                }
                if (partidaopen.Count > 0)
                {
                    ImpresionesDeDocumentosAutomaticas(partidaopen[0].ID_COMPROBANTE, "X");
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione un documento");
                    DocsAPagar.Clear();
                    ListViasPagos.Clear();
                }
                chkFiltro_UnChecked(chkFiltro, e);
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
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

        public void btnReimpr2_Click(object sender, RoutedEventArgs e)
        {

            List<DTE_SII> url = new List<DTE_SII>();
            if (chkDocFiscales.IsChecked == true)
            {
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            url.Add(DGDocCabec.Items.CurrentItem as DTE_SII);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                for (int h = 0; h < url.Count; h++)
                {
                    if (url[h].ISSELECTED == true)
                    {
                        DTE_SII partOpen = new DTE_SII();
                        partOpen.ISSELECTED = false;
                        partOpen.VBELN = url[h].VBELN;
                        partOpen.KONDA = url[h].KONDA;
                        partOpen.BUKRS = url[h].BUKRS;
                        partOpen.XBLNR = url[h].XBLNR;
                        partOpen.ZUONR = url[h].ZUONR;
                        partOpen.TDSII = url[h].TDSII;
                        partOpen.FODOC = url[h].FODOC;
                        partOpen.WAERS = url[h].WAERS;
                        partOpen.FECIMP = url[h].FECIMP;
                        partOpen.HORIM = url[h].HORIM;
                        partOpen.URLSII = url[h].URLSII;
                        url.Add(partOpen);
                    }
                }
                if (url.Count > 0)
                {
                    for (int i = 0; i < url.Count; i++)
                    {
                        if (url[i].ISSELECTED == true)
                        {
                            PDFViewer pdfvisor = new PDFViewer();
                            pdfvisor.Owner = this;
                            string url_reimpresion = "";
                            url_reimpresion = url[i].URLSII;
                            pdfvisor.webBrowser1.Navigate(url_reimpresion);
                            pdfvisor.Show();
                            url_reimpresion = "";
                        }
                    }
                }
            }
        }
        private void chkDocFiscales_Unchecked(object sender, RoutedEventArgs e)
        {
            RBDocReimp.IsChecked = false;
            txtComprReimp.Text = "";
            txtRUTReimp.Text = "";
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocCabec.Items.Clear();         
            btnReimpr.Visibility = Visibility.Visible;
            btnBuscarReimp.Content = "Buscar";
            GC.Collect();
        }
        private void chkDocFiscales_Checked(object sender, RoutedEventArgs e)
        {
            RBDocReimp.IsChecked = true;
            txtComprReimp.Text = "";
            txtRUTReimp.Text = "";
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocCabec.Items.Clear();
            btnReimpr.IsEnabled = false;
            btnReimpr2.IsEnabled = false;
            btnReimpr.Visibility = Visibility.Collapsed;
            btnBuscarReimp.Content = "Buscar";
            GC.Collect();
        }

        //Check que filtra la información en Anulaciones y Reimpresiones
        private void chkFiltro_Checked(object sender, RoutedEventArgs e)
        {
            if (GBReimpresion.IsVisible)
            {
                viaspagreimpraux.Clear();
                List<DOCUMENTOSAUX> DocsAPagar = new List<DOCUMENTOSAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as DOCUMENTOSAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<DOCUMENTOSAUX> partidaopen = new List<DOCUMENTOSAUX>();
                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        DOCUMENTOSAUX partOpen = new DOCUMENTOSAUX();
                        partOpen.ACC = DocsAPagar[k].ACC;
                        partOpen.CEBE = DocsAPagar[k].CEBE;
                        partOpen.CLASE_CUENTA = DocsAPagar[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CME = DocsAPagar[k].CME;
                        partOpen.APROBADOR_ANULA = DocsAPagar[k].APROBADOR_ANULA;
                        partOpen.APROBADOR_EX = DocsAPagar[k].APROBADOR_EX;
                        partOpen.CAJERO_GEN = DocsAPagar[k].CAJERO_GEN;
                        partOpen.CAJERO_RESP = DocsAPagar[k].CAJERO_RESP;
                        partOpen.CLIENTE = DocsAPagar[k].CLIENTE;
                        partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                        partOpen.EXCEPCION = DocsAPagar[k].EXCEPCION;
                        partOpen.FECHA_COMP = DocsAPagar[k].FECHA_COMP;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                        partOpen.FECHA_VENC_DOC = DocsAPagar[k].FECHA_VENC_DOC;
                        partOpen.HORA = DocsAPagar[k].HORA;
                        partOpen.MONTO_DIFERENCIA = DocsAPagar[k].MONTO_DIFERENCIA;
                        partOpen.MONTO_DOC = DocsAPagar[k].MONTO_DOC;
                        partOpen.NOTA_VENTA = DocsAPagar[k].NOTA_VENTA;
                        partOpen.NRO_ANULACION = DocsAPagar[k].NRO_ANULACION;
                        partOpen.NRO_COMPENSACION = DocsAPagar[k].NRO_COMPENSACION;
                        partOpen.NRO_DOCUMENTO = DocsAPagar[k].NRO_DOCUMENTO;
                        partOpen.NRO_REFERENCIA = DocsAPagar[k].NRO_REFERENCIA;
                        partOpen.SOCIEDAD = DocsAPagar[k].SOCIEDAD;
                        partOpen.NULO = DocsAPagar[k].NULO;
                        partOpen.NUM_CANCELACION = DocsAPagar[k].NUM_CANCELACION;
                        partOpen.NUM_CUOTA = DocsAPagar[k].NUM_CUOTA;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partOpen.PARCIAL = DocsAPagar[k].PARCIAL;
                        partOpen.POSICION = DocsAPagar[k].POSICION;
                        partOpen.TEXTO_CABECERA = DocsAPagar[k].TEXTO_CABECERA;
                        partOpen.TEXTO_EXCEPCION = DocsAPagar[k].TEXTO_EXCEPCION;
                        partOpen.TIME = DocsAPagar[k].TIME;
                        partOpen.TXT_ANULACION = DocsAPagar[k].TXT_ANULACION;
                        partOpen.USR_ANULADOR = DocsAPagar[k].USR_ANULADOR;
                        partidaopen.Add(partOpen);
                    }
                }
                if (partidaopen.Count > 0)
                {
                    List<VIAS_PAGO2> ListViasPagos = new List<VIAS_PAGO2>();
                    ListViasPagos.Clear();
                    for (int i = 1; i <= DGDocDet.Items.Count; i++)
                    {
                        if (i == 1)
                        {
                            DGDocDet.Items.MoveCurrentToFirst();
                        }
                        if (DGDocDet.Items.CurrentItem != null)
                        {
                            ListViasPagos.Add(DGDocDet.Items.CurrentItem as VIAS_PAGO2);
                        }

                        DGDocDet.Items.MoveCurrentToNext();
                    }
                    viaspagreimpr.Clear();
                    viaspagreimpraux.Clear();
                    for (int i = 0; i < ListViasPagos.Count; i++)
                    {
                        viaspagreimpr.Add(ListViasPagos[i]);
                        if (partidaopen[0].ID_COMPROBANTE == ListViasPagos[i].ID_COMPROBANTE)
                        {
                            viaspagreimpraux.Add(ListViasPagos[i]);
                        }
                    }
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = viaspagreimpraux;
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
            if (GBReimpresion.IsVisible)
            {
                if (viaspagreimpr.Count > 0)
                {
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = viaspagreimpr;
                    viaspagreimpraux.Clear();
                }
            }
            GC.Collect();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
         {
            CargarDatos();
            //GBInicio.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Visible;       
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            btnAutAnul.Visibility = Visibility.Collapsed;
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

            DGLogApertura.ItemsSource = null;
            DGLogApertura.Items.Clear();
            DGLogApertura.ItemsSource = logApertura;

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
            if (Reimpresio.IsMouseOver == true)
            {
                Reimp = new Reimpresion(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
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

        private void RBRutReimp_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTReimp.Text = "";
            txtRUTReimp.Visibility = Visibility.Visible;
            txtComprReimp.Text = "";
            txtComprReimp.Visibility = Visibility.Collapsed;
            docsreimpr = new List<DOCUMENTOS>();
            viaspagreimpr = new List<VIAS_PAGO2>();
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnBuscarReimp.IsEnabled = true;
            chkDocFiscales.IsChecked = false;
            GC.Collect();
        }

        private void RBDocReimp_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTReimp.Text = "";
            txtRUTReimp.Visibility = Visibility.Collapsed;
            txtComprReimp.Text = "";
            txtComprReimp.Visibility = Visibility.Visible;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            docsreimpr = new List<DOCUMENTOS>();
            viaspagreimpr = new List<VIAS_PAGO2>();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnBuscarReimp.IsEnabled = true;

            GC.Collect();
        }

        private void DGMonitor_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            timer.Stop();
            GC.Collect();
        }

        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {
            GC.Collect();
        }

        //BOTON QUE LLEVA EL DATO SELECCIONADO DESDE EL MONITOR 
        private void btnPagoMonitor_Click(object sender, RoutedEventArgs e)
        {
            GC.Collect();
        }
        private void ChkDetalleDocs_Checked(object sender, RoutedEventArgs e)
        {
            if (GBReimpresion.IsVisible == true)
            {
                List<DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<DOCUMENTOSAUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGDocCabec.Items.CurrentItem as DOCUMENTOSAUX);
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
            if (GBReimpresion.IsVisible == true)
            {
                List<DOCUMENTOSAUX> partidaseleccionadasaux2 = new List<DOCUMENTOSAUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGDocCabec.Items.CurrentItem as DOCUMENTOSAUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                for (int k = 0; k < partidaseleccionadasaux2.Count; k++)
                {
                    partidaseleccionadasaux2[k].ISSELECTED = false;
                }
                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();
                DGDocCabec.ItemsSource = partidaseleccionadasaux2;
            }
        }
        //CONEXION A LA RFC DEL MONITOR EN MODO MANUAL
        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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

        private void RecaudacionVeh_Click(object sender, RoutedEventArgs e)
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
