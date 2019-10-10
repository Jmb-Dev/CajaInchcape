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

namespace CajaIndigo.Vista.NotaCredito
{
    /// <summary>
    /// Interaction logic for NotaCredito.xaml
    /// </summary>
    public partial class NotaCredito : System.Windows.Window
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
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();
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
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.Reportes.Reportes Reporte;
        public NotaCredito()
        {
            InitializeComponent();
        }

        public NotaCredito(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
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
                PassUserCaja.Text = passlogg;
                idcaja.Text = cajaconect;
                NomCaja.Text = sucursal;
                SociedCaja.Text = sociedad;
                MonedaCaja.Text = moneda;
              
                cmbMoneda.Items.Clear();
                cmbMoneda.ItemsSource = myItemsCollection;
                if (cmbMoneda.SelectedValue != "0" && cmbMoneda.SelectedValue != "0")
                {
                    cmbMoneda.SelectedIndex = 0;
                }
                txtpais.Text = pais;
                lblPais.Content = pais;
                lblPassword.Content = passlogg;
                DateTime result = DateTime.Today;
                calendario.Text = Convert.ToString(result);
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

        private void btnBuscarNC_Click(object sender, RoutedEventArgs e)
        {
            if ((txtRUTNC.Text == "") && (txtComprNC.Text == ""))
            {
                System.Windows.MessageBox.Show("Ingrese un RUT o un número de comprobante");

            }
            else
            {
                ListaDocumentosNC();
            }
            //LimpiarViasDePago();
            chkFiltro.IsChecked = false;
            label10.Visibility = Visibility.Collapsed;
            DGDocDet.Visibility = Visibility.Collapsed;
            btnAutAnul.Visibility = Visibility.Collapsed;
            txtTotEfect.Visibility = Visibility.Collapsed;
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
            if (PagoDocumentos.IsMouseOver == true)
            {
                PagDocum = new PagoDocumento.PagoDocumento(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                PagDocum.Show();
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

        //FUNCION QUE TRAE LOS DOCUMENTOS PARA NOTAS DE CREDITO A PARTIR DE LA BUSQUEDA POR RUT O DOCUMENTO
        private void ListaDocumentosNC()
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

                if (RBRutNC.IsChecked == true)
                {
                    string RUTAux = "";
                    foreach (char value in txtRUTNC.Text.ToUpper())
                    {
                        if (value != 8207)
                        {
                            RUTAux = RUTAux + value;
                        }
                    }

                    RUTAux = RUTAux.Trim();

                    String RUT = DigitoVerificador(RUTAux);
                    //Verificacion de RUT
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRUTNC.Focus();
                    }
                    else
                    {
                        notasdecredito.notasdecredito(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprNC.Text, txtRUTNC.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "RUT", Convert.ToString(textBlock6.Content));
                    }
                }

                else if (RBDocNC.IsChecked == true)
                {
                    //***RFC Partidas abiertas para pago busqueda por numero de documento
                    string Documento = "";
                    Documento = txtComprNC.Text;
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
                        txtComprNC.Text = Documento;

                        notasdecredito.notasdecredito(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprNC.Text, txtRUTNC.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "Documento", Convert.ToString(textBlock6.Content));
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }

                if (notasdecredito.ObjDatosNC.Count > 0)
                {
                    GBDetalleDocs.Visibility = Visibility.Visible;
                    DGDocCabec.ItemsSource = null;
                    DGDocCabec.Items.Clear();
                    List<T_DOCUMENTOS_AUX> partidaopen = new List<T_DOCUMENTOS_AUX>();

                    for (int k = 0; k < notasdecredito.ObjDatosNC.Count; k++)
                    {
                        T_DOCUMENTOS_AUX partOpen = new T_DOCUMENTOS_AUX();
                        partOpen.ISSELECTED = false;
                        partOpen.ACC = notasdecredito.ObjDatosNC[k].ACC;
                        partOpen.CEBE = notasdecredito.ObjDatosNC[k].CEBE;
                        partOpen.CLASE_CUENTA = notasdecredito.ObjDatosNC[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = notasdecredito.ObjDatosNC[k].CLASE_DOC;
                        partOpen.CME = notasdecredito.ObjDatosNC[k].CME;
                        partOpen.COD_CLIENTE = notasdecredito.ObjDatosNC[k].COD_CLIENTE;
                        partOpen.COND_PAGO = notasdecredito.ObjDatosNC[k].COND_PAGO;
                        partOpen.CONTROL_CREDITO = notasdecredito.ObjDatosNC[k].CONTROL_CREDITO;
                        partOpen.DIAS_ATRASO = notasdecredito.ObjDatosNC[k].DIAS_ATRASO;
                        partOpen.ESTADO = notasdecredito.ObjDatosNC[k].ESTADO;
                        partOpen.FECHA_DOC = notasdecredito.ObjDatosNC[k].FECHA_DOC;
                        partOpen.FECVENCI = notasdecredito.ObjDatosNC[k].FECVENCI;
                        partOpen.ICONO = notasdecredito.ObjDatosNC[k].ICONO;
                        partOpen.MONEDA = notasdecredito.ObjDatosNC[k].MONEDA;
                        partOpen.MONTO = notasdecredito.ObjDatosNC[k].MONTO;
                        partOpen.MONTOF = notasdecredito.ObjDatosNC[k].MONTOF;
                        partOpen.MONTO_ABONADO = notasdecredito.ObjDatosNC[k].MONTO_ABONADO;
                        partOpen.MONTOF_ABON = notasdecredito.ObjDatosNC[k].MONTOF_ABON;
                        partOpen.MONTO_PAGAR = notasdecredito.ObjDatosNC[k].MONTO_PAGAR;
                        partOpen.MONTOF_PAGAR = notasdecredito.ObjDatosNC[k].MONTOF_PAGAR;
                        partOpen.NDOCTO = notasdecredito.ObjDatosNC[k].NDOCTO;
                        partOpen.NOMCLI = notasdecredito.ObjDatosNC[k].NOMCLI;
                        partOpen.NREF = notasdecredito.ObjDatosNC[k].NREF;
                        partOpen.RUTCLI = notasdecredito.ObjDatosNC[k].RUTCLI;
                        partOpen.SOCIEDAD = notasdecredito.ObjDatosNC[k].SOCIEDAD;
                        partOpen.BAPI = notasdecredito.ObjDatosNC[k].BAPI;
                        partOpen.FACT_ELECT = notasdecredito.ObjDatosNC[k].FACT_ELECT;
                        partOpen.FACT_SD_ORIGEN = notasdecredito.ObjDatosNC[k].FACT_SD_ORIGEN;
                        partOpen.ID_CAJA = notasdecredito.ObjDatosNC[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = notasdecredito.ObjDatosNC[k].ID_COMPROBANTE;
                        partOpen.LAND = notasdecredito.ObjDatosNC[k].LAND;
                        partidaopen.Add(partOpen);
                    }

                    if (partidaopen.Count > 0)
                    {
                        DGDocCabec.ItemsSource = partidaopen;
                        DGDocCabec.Visibility = Visibility.Visible;
                        DGDocDet.ItemsSource = null;
                        DGDocDet.Items.Clear();
                        DGDocDet.ItemsSource = notasdecredito.ViasPago;
                        DGDocDet.Visibility = Visibility.Collapsed;
                        label10.Visibility = Visibility.Collapsed;
                        btnRevisDoc.Visibility = Visibility.Visible;

                        if (chkNCTribut.IsChecked == true)
                        {
                            btnEmitirNC.IsEnabled = true;
                        }
                        else
                        {
                            btnEmitirNC.IsEnabled = false;
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

        private void btnEmitirNC_Click(object sender, RoutedEventArgs e)
        {
            List<T_DOCUMENTOS_AUX> DocsAPagar = new List<T_DOCUMENTOS_AUX>();
            DocsAPagar.Clear();
            if (this.DGDocCabec.Items.Count > 0)
            {
                for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                {
                    if (i == 0)
                        DGDocCabec.Items.MoveCurrentToFirst();
                    {
                        DocsAPagar.Add(DGDocCabec.Items.CurrentItem as T_DOCUMENTOS_AUX);
                    }
                    DGDocCabec.Items.MoveCurrentToNext();
                }
            }
            List<CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> partidaopen = new List<CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();

            for (int k = 0; k < DocsAPagar.Count; k++)
            {
                if (DocsAPagar[k].ISSELECTED == true)
                {
                    CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS partOpen = new CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS();
                    partOpen.ACC = DocsAPagar[k].ACC;
                    partOpen.CEBE = DocsAPagar[k].CEBE;
                    partOpen.CLASE_CUENTA = DocsAPagar[k].CLASE_CUENTA;
                    partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                    partOpen.CME = DocsAPagar[k].CME;
                    partOpen.COD_CLIENTE = DocsAPagar[k].COD_CLIENTE;
                    partOpen.COND_PAGO = DocsAPagar[k].COND_PAGO;
                    partOpen.CONTROL_CREDITO = DocsAPagar[k].CONTROL_CREDITO;
                    partOpen.DIAS_ATRASO = DocsAPagar[k].DIAS_ATRASO;
                    partOpen.ESTADO = DocsAPagar[k].ESTADO;
                    partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                    partOpen.FECVENCI = DocsAPagar[k].FECVENCI;
                    partOpen.ICONO = DocsAPagar[k].ICONO;
                    partOpen.MONEDA = DocsAPagar[k].MONEDA;
                    partOpen.MONTO = DocsAPagar[k].MONTO;
                    partOpen.MONTOF = DocsAPagar[k].MONTOF;
                    partOpen.MONTO_ABONADO = DocsAPagar[k].MONTO_ABONADO;
                    partOpen.MONTOF_ABON = DocsAPagar[k].MONTOF_ABON;
                    partOpen.MONTO_PAGAR = DocsAPagar[k].MONTO_PAGAR;
                    partOpen.MONTOF_PAGAR = DocsAPagar[k].MONTOF_PAGAR;
                    partOpen.NDOCTO = DocsAPagar[k].NDOCTO;
                    partOpen.NOMCLI = DocsAPagar[k].NOMCLI;
                    partOpen.NREF = DocsAPagar[k].NREF;
                    partOpen.RUTCLI = DocsAPagar[k].RUTCLI;
                    partOpen.SOCIEDAD = DocsAPagar[k].SOCIEDAD;
                    partOpen.BAPI = DocsAPagar[k].BAPI;
                    partOpen.FACT_ELECT = DocsAPagar[k].FACT_ELECT;
                    partOpen.FACT_SD_ORIGEN = DocsAPagar[k].FACT_SD_ORIGEN;
                    partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                    partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                    partOpen.LAND = DocsAPagar[k].LAND;
                    partidaopen.Add(partOpen);
                }
            }
            string checkTrib = "";
            if (chkNCTribut.IsChecked == true)
            {
                checkTrib = "X";
            }
                if (checkTrib != "X" && txtTotEfect.Text == "0")
                {

                    System.Windows.Forms.MessageBox.Show("Solo se puede Emitir Documento Tributario");
                }
                else
                {
                    NotasDeCreditoEmision NCEmitir = new NotasDeCreditoEmision();
                    NCEmitir.emitirnotasdecredito(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                        , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                        , Convert.ToString(cmbMoneda.SelectedItem), Convert.ToString(lblPais.Content), partidaopen, checkTrib);

                    if (NCEmitir.errormessage != "")
                    {
                        System.Windows.Forms.MessageBox.Show(NCEmitir.errormessage);
                        btnEmitirNC.IsEnabled = false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(NCEmitir.message);
                        label41.Visibility = Visibility.Collapsed;
                        txtTotEfect.Visibility = Visibility.Collapsed;
                        label10.Visibility = Visibility.Collapsed;
                        DGDocDet.Visibility = Visibility.Collapsed;
                        DGDocDet.ItemsSource = null;
                        DGDocDet.Items.Clear();
                        btnEmitirNC.IsEnabled = false;
                        ImpresionesDeDocumentosAutomaticas(NCEmitir.NumComprob, "X");
                    }
                }
            DGDocCabec.Visibility = Visibility.Collapsed;
            DGDocDet.Visibility = Visibility.Collapsed;
            label10.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            DGDocCabec.Visibility = Visibility.Collapsed;
            DGDocDet.Visibility = Visibility.Collapsed;
            label10.Visibility = Visibility.Collapsed;
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
        //Check que filtra la información en Anulaciones y Reimpresiones
        private void chkFiltro_Checked(object sender, RoutedEventArgs e)
        {
            if (GBEmisionNC.IsVisible)
            {
                viaspagreimpraux.Clear();
                List<T_DOCUMENTOS_AUX> DocsAPagar = new List<T_DOCUMENTOS_AUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as T_DOCUMENTOS_AUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> partidaopen = new List<CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS partOpen = new CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS();
                        partOpen.ACC = DocsAPagar[k].ACC;
                        partOpen.CEBE = DocsAPagar[k].CEBE;
                        partOpen.CLASE_CUENTA = DocsAPagar[k].CLASE_CUENTA;
                        partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                        partOpen.CME = DocsAPagar[k].CME;
                        partOpen.COD_CLIENTE = DocsAPagar[k].COD_CLIENTE;
                        partOpen.COND_PAGO = DocsAPagar[k].COND_PAGO;
                        partOpen.CONTROL_CREDITO = DocsAPagar[k].CONTROL_CREDITO;
                        partOpen.DIAS_ATRASO = DocsAPagar[k].DIAS_ATRASO;
                        partOpen.ESTADO = DocsAPagar[k].ESTADO;
                        partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                        partOpen.FECVENCI = DocsAPagar[k].FECVENCI;
                        partOpen.ICONO = DocsAPagar[k].ICONO;
                        partOpen.MONEDA = DocsAPagar[k].MONEDA;
                        partOpen.MONTO = DocsAPagar[k].MONTO;
                        partOpen.MONTOF = DocsAPagar[k].MONTOF;
                        partOpen.MONTO_ABONADO = DocsAPagar[k].MONTO_ABONADO;
                        partOpen.MONTOF_ABON = DocsAPagar[k].MONTOF_ABON;
                        partOpen.MONTO_PAGAR = DocsAPagar[k].MONTO_PAGAR;
                        partOpen.MONTOF_PAGAR = DocsAPagar[k].MONTOF_PAGAR;
                        partOpen.NDOCTO = DocsAPagar[k].NDOCTO;
                        partOpen.NOMCLI = DocsAPagar[k].NOMCLI;
                        partOpen.NREF = DocsAPagar[k].NREF;
                        partOpen.RUTCLI = DocsAPagar[k].RUTCLI;
                        partOpen.SOCIEDAD = DocsAPagar[k].SOCIEDAD;
                        partOpen.BAPI = DocsAPagar[k].BAPI;
                        partOpen.FACT_ELECT = DocsAPagar[k].FACT_ELECT;
                        partOpen.FACT_SD_ORIGEN = DocsAPagar[k].FACT_SD_ORIGEN;
                        partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                        partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                        partOpen.LAND = DocsAPagar[k].LAND;
                        partidaopen.Add(partOpen);
                    }
                }
                if (partidaopen.Count > 0)
                {
                    List<VIAS_PAGO2> ListViasPagos = new List<VIAS_PAGO2>();

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
                    DGDocDet.Visibility = Visibility.Visible;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Seleccione un registro en la tabla de documentos/cabecera");
                }
            }

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
            if (GBEmisionNC.IsVisible)
            {
                if (viaspagreimpr.Count > 0)
                {
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = viaspagreimpr;
                    viaspagreimpraux.Clear();
                }
                DGDocDet.Visibility = Visibility.Collapsed;
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

        private void cmbBancoProp_DropDownClosed(object sender, EventArgs e)
        {
            int posicion;
            posicion = cmbBancoProp.SelectedIndex;
            cmbCuentasBancosProp.SelectedIndex = posicion;
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
                                //label41.Visibility = Visibility.Collapsed;
                                //txtCodAuto.Visibility = Visibility.Collapsed;
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
                                //label41.Visibility = Visibility.Visible;
                                //txtCodAuto.Visibility = Visibility.Visible;
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

        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {
            GC.Collect();
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

        private void RBDocNC_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTNC.Text = "";
            txtRUTNC.Visibility = Visibility.Collapsed;
            txtComprNC.Text = "";
            txtComprNC.Visibility = Visibility.Visible;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            viaspagreimpr = new List<VIAS_PAGO2>();
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnEmitirNC.IsEnabled = false;
            chkNCTribut.IsChecked = false;
            chkNCTribut.IsEnabled = true;
            GC.Collect();
        }

        private void RBRutNC_Checked(object sender, RoutedEventArgs e)
        {
            txtRUTNC.Text = "";
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            txtRUTNC.Visibility = Visibility.Visible;
            txtComprNC.Text = "";
            txtComprNC.Visibility = Visibility.Collapsed;
            viaspagreimpr = new List<VIAS_PAGO2>();
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            btnEmitirNC.IsEnabled = false;
            chkNCTribut.IsChecked = false;
            chkNCTribut.IsEnabled = true;
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
                cmbIfinan.ItemsSource = listabancos;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos de instituciones financieras en el sistema");
            }
            GC.Collect();
        }
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

        private void btnRevisDoc_Click(object sender, RoutedEventArgs e)
        {
            List<T_DOCUMENTOS_AUX> DocsAPagar = new List<T_DOCUMENTOS_AUX>();
            DocsAPagar.Clear();
            if (this.DGDocCabec.Items.Count > 0)
            {
                for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                {
                    if (i == 0)
                        DGDocCabec.Items.MoveCurrentToFirst();
                    {
                        DocsAPagar.Add(DGDocCabec.Items.CurrentItem as T_DOCUMENTOS_AUX);
                    }
                    DGDocCabec.Items.MoveCurrentToNext();
                }
            }
            List<CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> partidaopen = new List<CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();

            for (int k = 0; k < DocsAPagar.Count; k++)
            {
                if (DocsAPagar[k].ISSELECTED == true)
                {
                    CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS partOpen = new CajaIndigo.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS();
                    partOpen.ACC = DocsAPagar[k].ACC;
                    partOpen.CEBE = DocsAPagar[k].CEBE;
                    partOpen.CLASE_CUENTA = DocsAPagar[k].CLASE_CUENTA;
                    partOpen.CLASE_DOC = DocsAPagar[k].CLASE_DOC;
                    partOpen.CME = DocsAPagar[k].CME;
                    partOpen.COD_CLIENTE = DocsAPagar[k].COD_CLIENTE;
                    partOpen.COND_PAGO = DocsAPagar[k].COND_PAGO;
                    partOpen.CONTROL_CREDITO = DocsAPagar[k].CONTROL_CREDITO;
                    partOpen.DIAS_ATRASO = DocsAPagar[k].DIAS_ATRASO;
                    partOpen.ESTADO = DocsAPagar[k].ESTADO;
                    partOpen.FECHA_DOC = DocsAPagar[k].FECHA_DOC;
                    partOpen.FECVENCI = DocsAPagar[k].FECVENCI;
                    partOpen.ICONO = DocsAPagar[k].ICONO;
                    partOpen.MONEDA = DocsAPagar[k].MONEDA;
                    partOpen.MONTO = DocsAPagar[k].MONTO;
                    partOpen.MONTOF = DocsAPagar[k].MONTOF;
                    partOpen.MONTO_ABONADO = DocsAPagar[k].MONTO_ABONADO;
                    partOpen.MONTOF_ABON = DocsAPagar[k].MONTOF_ABON;
                    partOpen.MONTO_PAGAR = DocsAPagar[k].MONTO_PAGAR;
                    partOpen.MONTOF_PAGAR = DocsAPagar[k].MONTOF_PAGAR;
                    partOpen.NDOCTO = DocsAPagar[k].NDOCTO;
                    partOpen.NOMCLI = DocsAPagar[k].NOMCLI;
                    partOpen.NREF = DocsAPagar[k].NREF;
                    partOpen.RUTCLI = DocsAPagar[k].RUTCLI;
                    partOpen.SOCIEDAD = DocsAPagar[k].SOCIEDAD;
                    partOpen.BAPI = DocsAPagar[k].BAPI;
                    partOpen.FACT_ELECT = DocsAPagar[k].FACT_ELECT;
                    partOpen.FACT_SD_ORIGEN = DocsAPagar[k].FACT_SD_ORIGEN;
                    partOpen.ID_CAJA = DocsAPagar[k].ID_CAJA;
                    partOpen.ID_COMPROBANTE = DocsAPagar[k].ID_COMPROBANTE;
                    partOpen.LAND = DocsAPagar[k].LAND;
                    partidaopen.Add(partOpen);
                }
            }
            NotasDeCreditoCheck NCCheck = new NotasDeCreditoCheck();
            NCCheck.chequearnotascreditos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(cmbMoneda.SelectedItem), Convert.ToString(lblPais.Content), partidaopen);
            if (NCCheck.Efectivo != "00")
            {
                txtTotEfect.Text = NCCheck.Efectivo;
            }
            else
            {
                txtTotEfect.Text = "0";
            }
            string Mensaje = NCCheck.errormessage;

            if (NCCheck.errormessage != "")
            {
                label41.Visibility = Visibility.Visible;
                txtTotEfect.Visibility = Visibility.Visible;
                label10.Visibility = Visibility.Visible;
                DGDocDet.Visibility = Visibility.Visible;
                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();
                DGDocDet.ItemsSource = NCCheck.viapago;
                btnEmitirNC.IsEnabled = true;
                chkNCTribut.IsChecked = true;
                chkNCTribut.IsEnabled = true;
                label41.Visibility = Visibility.Visible;
                txtTotEfect.Visibility = Visibility.Visible;
                System.Windows.Forms.MessageBox.Show("Solo se puede emitir Documento Tributario");
            }
            else
            {
                label41.Visibility = Visibility.Visible;
                txtTotEfect.Visibility = Visibility.Visible;
                label10.Visibility = Visibility.Visible;
                DGDocDet.Visibility = Visibility.Visible;
                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();
                DGDocDet.ItemsSource = NCCheck.viapago;
                btnEmitirNC.IsEnabled = true;
                label41.Visibility = Visibility.Visible;
                txtTotEfect.Visibility = Visibility.Visible;
            }
        }
        //if (txtTotEfect.Text == "0")
        //if( NCCheck.errormessage != "")
        // {

        //     label41.Visibility = Visibility.Visible;
        //     txtTotEfect.Visibility = Visibility.Visible;
        //     label10.Visibility = Visibility.Visible;
        //     DGDocDet.Visibility = Visibility.Visible;
        //     DGDocDet.ItemsSource = null;
        //     DGDocDet.Items.Clear();
        //     DGDocDet.ItemsSource = NCCheck.viapago;
        //     btnEmitirNC.IsEnabled = true;
        //     chkNCTribut.IsChecked = true;
        //     chkNCTribut.IsEnabled = true;
        //     label41.Visibility = Visibility.Visible;
        //     txtTotEfect.Visibility = Visibility.Visible;
        //     System.Windows.Forms.MessageBox.Show("Solo se puede emitir Documento Tributario");
        // }
        // else
        // {
        //     label41.Visibility = Visibility.Visible;
        //     txtTotEfect.Visibility = Visibility.Visible;
        //     label10.Visibility = Visibility.Visible;
        //     DGDocDet.Visibility = Visibility.Visible;
        //     DGDocDet.ItemsSource = null;
        //     DGDocDet.Items.Clear();
        //     DGDocDet.ItemsSource = NCCheck.viapago;
        //     btnEmitirNC.IsEnabled = true;
        //     label41.Visibility = Visibility.Visible;
        //     txtTotEfect.Visibility = Visibility.Visible;
        // }
        //if (NCCheck.errormessage != "")
        //{
        //    System.Windows.Forms.MessageBox.Show(NCCheck.errormessage);
        //    btnEmitirNC.IsEnabled = true;
        //    chkNCTribut.IsChecked = true;
        //    chkNCTribut.IsEnabled = true;
        //}
        //else
        //{
        //    if (NCCheck.message != "")
        //    {
        //        System.Windows.Forms.MessageBox.Show(NCCheck.message);
        //    }
        //    label41.Visibility = Visibility.Visible;
        //    txtTotEfect.Visibility = Visibility.Visible;
        //    label10.Visibility = Visibility.Visible;
        //    DGDocDet.Visibility = Visibility.Visible;
        //    DGDocDet.ItemsSource = null;
        //    DGDocDet.Items.Clear();
        //    DGDocDet.ItemsSource = NCCheck.viapago;
        //    btnEmitirNC.IsEnabled = true;
        //}
        //GC.Collect();

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
    }
}
