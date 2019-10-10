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
//using CajaIndu.PDFPageNumber;
using CajaIndu.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndu.AppPersistencia.Class.PartidasAbiertas;
using CajaIndu.AppPersistencia.Class.Monitor.Estructura;
using CajaIndu.AppPersistencia.Class.Monitor;
using CajaIndu.AppPersistencia.Class.MatrizDePago.Estructura;
using CajaIndu.AppPersistencia.Class.MatrizDePago;
using CajaIndu.AppPersistencia.Class.DocumentosPagosMasivos;
using CajaIndu.AppPersistencia.Class.DocumentosPagosMasivos.Estructura;
using CajaIndu.AppPersistencia.Class.StatusPagosChq;
using CajaIndu.AppPersistencia.Class.StatusPagosChq.Estructura;
using CajaIndu.AppPersistencia.Class.PagoDocumentosIngreso.Estructura;
using CajaIndu.AppPersistencia.Class.PagoDocumentosIngreso;
using CajaIndu.AppPersistencia.Class.MaestroDeBancos.Estructura;
using CajaIndu.AppPersistencia.Class.MaestroDeBancos;
using CajaIndu.AppPersistencia.Class.MaestroFinancieras.Estructura;
using CajaIndu.AppPersistencia.Class.MaestroFinancieras;
using CajaIndu.AppPersistencia.Class.Login;
using CajaIndu.AppPersistencia.Class.Login.Estructura;
using CajaIndu.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndu.AppPersistencia.Class.CierreCaja;
using CajaIndu.AppPersistencia.Class.ArqueoCaja.Estructura;
using CajaIndu.AppPersistencia.Class.ArqueoCaja;
using CajaIndu.AppPersistencia.Class.PreCierreCaja.Estructura;
using CajaIndu.AppPersistencia.Class.PreCierreCaja;
using CajaIndu.AppPersistencia.Class.CierreCajaDefinitvo.Estructura;
using CajaIndu.AppPersistencia.Class.CierreCajaDefinitvo;
using CajaIndu.AppPersistencia.Class.Anticipos;
using CajaIndu.AppPersistencia.Class.PagoAnticipos;
using CajaIndu.AppPersistencia.Class.PagoAnticipos.Estructura;
using CajaIndu.AppPersistencia.Class.BusquedaAnulacion.Estructura;
using CajaIndu.AppPersistencia.Class.BusquedaAnulacion;
using CajaIndu.AppPersistencia.Class.BusquedaReimpresiones.Estructura;
using CajaIndu.AppPersistencia.Class.BusquedaReimpresiones;
using CajaIndu.AppPersistencia.Class.NotasDeCredito;
using CajaIndu.AppPersistencia.Class.AnulacionComprobantes;
using CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura;
using CajaIndu.AppPersistencia.Class.UsuariosCaja;
using CajaIndu.AppPersistencia.Class.BloquearCaja;
using CajaIndu.AppPersistencia.Class.ReimpresionFiscal;
using CajaIndu.AppPersistencia.Class.ReimpresionComprobantes;
using CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura;
using CajaIndu.AppPersistencia.Class.MaestroTarjetas;
using CajaIndu.AppPersistencia.Class.MaestroTarjetas.Estructura;
using CajaIndu.AppPersistencia.Class.RendicionCaja;
using CajaIndu.AppPersistencia.Class.RendicionCaja.Estructura;
using CajaIndu.AppPersistencia.Class.RecaudacionVehiculos.Estructura;
using CajaIndu.AppPersistencia.Class.RecaudacionVehiculos;
using CajaIndu.AppPersistencia.Class.NotasDeCreditoCheck.Estructura;
using CajaIndu.AppPersistencia.Class.NotasDeCreditoCheck;
using CajaIndu.AppPersistencia.Class.NotasDeCreditoEmision;
using CajaIndu.AppPersistencia.Class.GestionDeDepositos;
using CajaIndu.AppPersistencia.Class.GestionDeDepositos.Estructura;
using CajaIndu.AppPersistencia.Class.CheckUserAnulacion;
using CajaIndu.AppPersistencia.Class.DepositoProceso;
using CajaIndu.AppPersistencia.Class.ReportesCaja.Estructura;
using CajaIndu.AppPersistencia.Class.ReportesCaja;
using CajaIndu.AppPersistencia.Class.PagosMasivosNew;
using CajaIndu.AppPersistencia.Class.ReimpresionFiscal.Estructura;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Interop.Excel;
using CajaIndu;
using System.Text.RegularExpressions;



namespace CajaIndu 
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    ///  private PdfTemplate totalPages;
    
    public partial class PagosDocumentos : System.Windows.Window
    {
         private PdfTemplate totalPages;
         private PdfWriter Write;
        
 //       List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
        List<DetalleViasPago> cheques = new List<DetalleViasPago>();
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
        List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.VIAS_PAGO> viaspagreimprcompr = new List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.VIAS_PAGO>();
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

        public PagosDocumentos(string usuariologg, string passlogg,string usuariotemp, string cajaconect, string sucursal, string sociedad, List<string> moneda, string pais, double monto, List<LOG_APERTURA> logapertura)
        {
            try
            {
                InitializeComponent();
                
                if (moneda.Count == 0) 
                {
                    moneda.Add("CLP");
                    moneda.Add("USS");
                }
                int test = 0;
                test = cmbMoneda.Items.Count;
                //GroupBox GBInicio = new GroupBox();
                GBInicio.Visibility = Visibility.Visible;
                GBMonitor.Visibility = Visibility.Visible;
                GBPagoDocs.Visibility = Visibility.Collapsed;
                GBPagoDocs.Visibility = Visibility.Collapsed;
                GBInicio.Visibility = Visibility.Collapsed;
                GBAnulacion.Visibility = Visibility.Collapsed;
                GBReimpresion.Visibility = Visibility.Collapsed;
                GBViasPago.Visibility = Visibility.Collapsed;
                GBDocsAPagar.Visibility = Visibility.Collapsed;
                GBDetalleDocs.Visibility = Visibility.Collapsed;
                GBEmisionNC.Visibility = Visibility.Collapsed;
                GBRendicion.Visibility = Visibility.Collapsed;
                GBrecauda.Visibility = Visibility.Collapsed;
                GBDocs.Visibility = Visibility.Collapsed;
                GBResumenCaja.Visibility = Visibility.Collapsed;
                GBDetEfectivo.Visibility = Visibility.Collapsed;
                GBCierreCaja.Visibility = Visibility.Collapsed;
                GBCommentCierre.Visibility = Visibility.Collapsed;
                GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
                GBGestionDeBancos.Visibility = Visibility.Collapsed;
                GBReportes.Visibility = Visibility.Collapsed;
                textBlock6.Content = cajaconect;
                textBlock7.Content = usuariologg;
                textBlock8.Content = sucursal;
                textBlock9.Content = usuariotemp ;
                lblMonto.Content = Convert.ToString(monto);
                lblSociedad.Content = sociedad;
                cmbMoneda.Items.Clear();
                cmbMoneda.ItemsSource = moneda;
                if (moneda.Count == 1)
                {
                    cmbMoneda.SelectedIndex = 0;
                }

                lblPais.Content = pais;
                lblPassword.Content = passlogg;
                DateTime result = DateTime.Today;
                datePicker1.Text = Convert.ToString(result);

                DGLogApertura.ItemsSource = null;
                DGLogApertura.Items.Clear();
                DGLogApertura.ItemsSource = logapertura;
              
                //maestrobancos.maestrobancos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), Convert.ToString(lblPais.Content), Convert.ToString(cmbMoneda.Text),Convert.ToString(lblSociedad.Content));
                //if (maestrobancos.T_Retorno.Count > 0)
                //{
                //    cmbBanco.ItemsSource = null;
                //    cmbBanco.Items.Clear();
                //    List<string> listabancos = new List<string>();

                //    for (int i = 0; i < maestrobancos.T_Retorno.Count; i++)
                //    {
                //        listabancos.Add(maestrobancos.T_Retorno[i].BANKL + " - " + maestrobancos.T_Retorno[i].BANKA);
                //    }
                //    //cmbBanco.ItemsSource = maestrobancos.T_Retorno[0].BANKL + " - " + maestrobancos.T_Retorno[0].BANKA;
                //    cmbBanco.ItemsSource = listabancos;
                //}
                //else
                //{
                //    System.Windows.Forms.MessageBox.Show("No existen datos de bancos en el sistema");
                //}
                GC.Collect();  
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content),Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                
            }          
        }


        private void Window_Loaded()
        {
            GBInicio.Visibility = Visibility.Visible;
            GBMonitor.Visibility = Visibility.Visible;
            //throw new NotImplementedException();
            DateTime result = DateTime.Today;
            datePicker1.Text = Convert.ToString(result);
            //ACTIVACION DEL MONITOR
            chkMonitor.IsChecked = true;

            //if (chkMonitor.IsChecked.Value)
            //{
            //    timer.Interval = TimeSpan.FromSeconds(15);
            //    timer.Tick += timer_Tick;
            //    timer.Start();
            //}
           
            //RFC PARA OBTENER LOS BANCOS
            RFC_Combo_Bancos();
            GC.Collect();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //MainWindow frm = new MainWindow();
            //frm.Visibility = Visibility.Visible;
            //frm.Show();
            
            timer.Stop();
            MainWindow window = System.Windows.Window.GetWindow(this.Owner) as MainWindow;
            if (window != null)
            {
                this.Close();
                window.Visibility = Visibility.Visible;
            }
           // Process proc = new Process();
            //proc.Kill("CajaLogin"); //GetProcessesByName("CajaLogin"); 
	        //proc[0].Kill();
            //Process myProcess;
            //myProcess = Process.Start("CajaLogin");

            //myProcess.Kill();
        }



        //MANEJO DE LOS EVENTOS ASOCIADOS A TABCONTROLS
        #region TabControl
        //EVENTO DE TAB CONTROL CUANDO ESTE CAMBIA EN EL PAGO DE DOCUMENTOS
        private void tabControlReca_SelectionChanged(object sender, SelectionChangedEventArgs e)
        { 
            
 
        }


        private void tabControlAnulacion_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            txtDocuAnt.Text = "";
            txtComprAnV.Text = "";
            txtComprAn.Text = "";
            txtRUTAnt.Text = "";
            txtRUTAnV.Text = "";
            txtRUTAn.Text = "";
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
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
 
        //MANEJO DE LOS EVENTOS ASOCIADOS A LOS CHECKBOXS
        #region CheckBox's
        //Check que activa desactiva manualmente el Monitor
        private void chkMonitor_Checked(object sender, RoutedEventArgs e)
        {
          //  timer.Start();
        }
        private void chkMonitor_UnChecked(object sender, RoutedEventArgs e)
        {
           // timer.Stop();
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
                    for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as T_DOCUMENTOS_AUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> partidaopen = new List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS partOpen = new CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS();
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
                // viaspagreimpr.Clear();
                //List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> DocsAPagar = new List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();
                //for (int i = 0; i < DGDocCabec.SelectedItems.Count; i++)
                //{
                //    {
                //        DocsAPagar.Add(DGDocCabec.SelectedItems[i] as CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS);
                //    }
                //}
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
                    for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
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
                         //IdComprobante = DocsAPagar[k].ID_COMPROBANTE;
                         partidaopen.Add(partOpen);
                     }
                 }
                //detalle.Clear();
                //List<CAB_COMP> CabeceraDocs = new List<CAB_COMP>();
                //for (int i = 0; i < DGDocCabec.SelectedItems.Count; i++)
                //{
                //    {
                //        CabeceraDocs.Add(DGDocCabec.SelectedItems[i] as CAB_COMP);
                //    }
                //}
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

            if (GBReimpresion.IsVisible)
            {
                viaspagreimpraux.Clear();
                // viaspagreimpr.Clear();
                List<DOCUMENTOSAUX> DocsAPagar = new List<DOCUMENTOSAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
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
                //List<DOCUMENTOS> DocsAPagar = new List<DOCUMENTOS>();
                //for (int i = 0; i < DGDocCabec.SelectedItems.Count; i++)
                //{
                //    {
                //        DocsAPagar.Add(DGDocCabec.SelectedItems[i] as DOCUMENTOS);
                //    }
                //}
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
            if (GBAnulacion.IsVisible)
            {
                if (detalle.Count > 0)
                { 
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = detalle;
                    detalleaux.Clear();
                    //detalle.Clear();
                }
            }
            if (GBReimpresion.IsVisible)
            {
                if (viaspagreimpr.Count > 0)
                {
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = viaspagreimpr;
                    viaspagreimpraux.Clear();
                    //viaspagreimpr.Clear();
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
                    //viaspagreimpr.Clear();
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
            //partidaseleccionadasaux2.Clear();
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
            //partidaseleccionadasaux2.Clear();
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
            //partidaseleccionadasaux2.Clear();
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
            //partidaseleccionadasaux2.Clear();
            GC.Collect();
        }




        private void DataGridCheckBoxColumn_Checked(object sender, RoutedEventArgs e)
        {

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
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as CAB_COMPAUX);
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
            if (GBEmisionNC.IsVisible == true)
            {
                List<T_DOCUMENTOS_AUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOS_AUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as T_DOCUMENTOS_AUX);
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
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as DOCUMENTOSAUX);
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
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as CAB_COMPAUX);
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
            if (GBEmisionNC.IsVisible == true)
            {
                List<T_DOCUMENTOS_AUX> partidaseleccionadasaux2 = new List<T_DOCUMENTOS_AUX>();
                partidaseleccionadasaux2.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count - 1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as T_DOCUMENTOS_AUX);
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
                            partidaseleccionadasaux2.Add(DGRecau.Items.CurrentItem as DOCUMENTOSAUX);
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
            //partidaseleccionadasaux2.Clear();
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
            //partidaseleccionadasaux2.Clear();
            GC.Collect();
        }

        private void chkAbono_Checked(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        //MANEJO DE LOS EVENTOS ASOCIADOS A LOS BOTONES
        #region Botones

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
                    //MessageBox.Show("Conectandose a la RFC del Monitor en modo manual");
                }
                GC.Collect();
            }
            catch  (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                
            }                 
        }
                    
        
        //CLICK BOTON DE BARRA DE HERRAMIENTAS QUE ACTIVA EL PAGO DE DOCUMENTOS
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
          
            GBPagoDocs.Visibility = Visibility.Visible;
            GBInicio.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
			GBrecauda.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            //LimpiarCamposInformeRendicion();
            GC.Collect();
        }


        //CLICK BOTON DE BARRA DE HERRAMIENTAS QUE ACTIVA LA ANULACION DE DOCUMENTOS
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Visible;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
			GBrecauda.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            //btnAutAnul.Visibility = Visibility.Visible;
            //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            //LimpiarCamposInformeRendicion();
            GC.Collect();
        }


        //CLICK BOTON DE MENU QUE ACTIVA LA REIMPRESION DE DOCUMENTOS
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            //GBAnticipos.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Visible;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
			GBrecauda.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            btnAutAnul.Visibility = Visibility.Collapsed;
            //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
        
            //LimpiarCamposInformeRendicion();
            GC.Collect();
        }


        //BOTON DE MENU PARA EL CIERRE DE CAJA
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            //GBAnticipos.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
			GBrecauda.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBMonitor.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Visible;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            LimpiarElementosDeCierreDeCaja();
           // LimpiarCamposInformeRendicion();
            LimpiarViasDePago();
            LimpiarEntradasDeDatos();

            //DPickDesde.Text = "";// Convert.ToString(datePicker1.SelectedDate);
            //DPickHasta.Text = "";// Convert.ToString(datePicker1.SelectedDate);
            //GBResumenCaja.Visibility = Visibility.Visible;
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
            if (LogApert[0].MONEDA == "CLP")
            {
                if (LogApert[0].MONTO.Contains(","))
                {
                    LogApert[0].MONTO = LogApert[0].MONTO.Replace(",", "");
                    LogApert[0].MONTO = LogApert[0].MONTO.Substring(0, LogApert[0].MONTO.Length - 2);
                }
                txtMontoApert.Text = LogApert[0].MONTO;
            }
            else
            {
                txtMontoApert.Text = LogApert[0].MONTO;
            }
            GC.Collect();
         }
        //MENU RECAUDACION DE VEHICULO
        private void bt_recaudacion(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBrecauda.Visibility = Visibility.Visible;
            GBDocs.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            LimpiarEntradasDeDatos();
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarCamposInformeRendicion();
            GC.Collect();
        }
        //MENU GESTION DE DEPOSITOS
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBrecauda.Visibility = Visibility.Visible;
            GBDocs.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            LimpiarEntradasDeDatos();
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarCamposInformeRendicion();
            GC.Collect();
        }

        //RECAUDACION DE VEHICULOS
        private void btnPago_Click(object sender, RoutedEventArgs e)
        {
            //DGRecau.SelectedItem
            bool daleclavo = true;
            List<AutorizacionViasPago> DocsAPagar = new List<AutorizacionViasPago>();
            if (DGAutorizacionVehiculos.Items.Count > 0)
            {
                DGAutorizacionVehiculos.SelectAll();
            }
                daleclavo = false;
                //for (int i = 0; i < DGAutorizacionVehiculos.SelectedItems.Count; i++)
                //{
                //    {
                //        DocsAPagar.Add(DGAutorizacionVehiculos.SelectedItems[i] as AutorizacionViasPago);
                //    }
                //}
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
                            daleclavo = true;
                            // RecaudacionVehiculos(DocsAPagar);
                        }
                        else
                        {
                            daleclavo = false;
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
                        daleclavo = true;

                    }
                    else
                    {
                        daleclavo = false;
                        validador = validador + 1;
                        System.Windows.Forms.MessageBox.Show("Ingrese el número de documento de vía de pago en documento N°" + DocsAPagar[j].VBELN);
                    }

                }
                if ((DocsAPagar[j].VIADP == "U") | (DocsAPagar[j].VIADP == "B"))
                {

                    if (DocsAPagar[j].FEC_EMISION != "")
                    {
                        daleclavo = true;
                       
                    }
                    else
                    {
                        daleclavo = false;
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
                            if (DocsAPagar[j].AUTORIZACION != "")  //& (DocsAPagar[j].ASIGNACION != ""))
                            {
                               // RecaudacionVehiculos(DocsAPagar);
                                daleclavo = true;
                            }
                            else
                            {
                                daleclavo = false;
                                validador = validador + 1;
                                System.Windows.Forms.MessageBox.Show("Ingrese el código de autorización de la tarjeta de crédito ó débito en documento N° " + DocsAPagar[j].VBELN);
                       
                            }
                        }
                        else
                        {
                            validador = validador + 1;
                            System.Windows.Forms.MessageBox.Show("Ingrese el código de operación de la tarjeta de crédito ó débito en documento N° " + DocsAPagar[j].VBELN);
                       
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

            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
                
           
            if (validador < 1)
            {
                RecaudacionVehiculos(DocsAPagar, LogApert[0].ID_REGISTRO);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No se pudo procesar la recaudación de vehículos");
            }
            GC.Collect();
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
                if (chkAbono.IsChecked == true)
                {
                    for (int i = 0; i < partidaseleccionadas.Count; i++)
                    {
                        if (partidaseleccionadas[i].NREF != "")
                        {
                            if (partidaseleccionadas[i].NREF.Substring(0, 1) != "E")
                            {
                                System.Windows.Forms.MessageBox.Show("No se puede realizar abonos parciales a documentos que no sean partidas abiertas:" + partidaseleccionadas[i].NDOCTO);
                                permiso = false;
                            }
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("No se puede realizar abonos parciales a documentos que no sean partidas abiertas:" + partidaseleccionadas[i].NDOCTO);
                            permiso = false;
                            
                        }
                    }

                
                }


                //if (this.DGPagos.SelectedItems.Count > 0)
                //    for (int i = 0; i < DGPagos.SelectedItems.Count; i++)
                //    {
                //        {
                //            partidaseleccionadas.Add(DGPagos.SelectedItems[i] as T_DOCUMENTOS);
                //        }
                //    }
                if (permiso == true)
                {
                    List<ViasPago> Condiciones = new List<ViasPago>();
                    List<string> CondicionPago = new List<string>();
                    string CondPago = "";

                    ViasPago Condic;// = new ViasPago(acc, cond_pago, caja);
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
                            // if (CondicionPago.Count < 2)
                            // {
                            Condic = new ViasPago(partidaseleccionadas[i].ACC, partidaseleccionadas[i].COND_PAGO, partidaseleccionadas[i].CME);
                            Condic = new ViasPago(partidaseleccionadas[i].ACC, partidaseleccionadas[i].COND_PAGO, Convert.ToString(textBlock6.Content));
                            Condiciones.Add(Condic);
                            //}
                            //else
                            //{
                            //    System.Windows.MessageBox.Show("Registros con dos o mas condiciones de pago distintas");
                            //    break;

                            //}

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message + ex.StackTrace);
                            System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                        }
                    }
                    //MODIFICACION A CONDICIONES DE PAGO DISTINTAS
                    //if (CondicionPago.Count < 2)
                    //{
                    //RFC para consulta de estatus de cobro del cliente selecccionado
                    //partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtRut.Text,"", txtRut.Text, Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text));
                    DateTime Anual = datePicker1.SelectedDate.Value;
                    String EjercicioValue = Convert.ToString(Anual.Year);
                    //EstatusCobranza estatuscobranza = new EstatusCobranza();
                    //estatuscobranza.EstatusCobro(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), Convert.ToString(lblSociedad.Content), txtRut.Text, "09", "M", "W", EjercicioValue);
                    //if (estatuscobranza.protestado == "")
                    //{
                    //    //estatuscobranza.protestado = "X";
                    //}
                    //else
                    //{
                    //    estatuscobranza.protestado = "X";
                    //    System.Windows.MessageBox.Show("Este cliente presenta cheque(s) protestado(s). " + estatuscobranza.message);
                    //}
                    //RFC que retorna las formas de pago de acuerdo a los registros seleccionados
                    MatrizDePago matrizpago = new MatrizDePago();
                    // List<ViasPago> LVP = new List<ViasPago>();
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
                        if(Excepcion2 == true & Monto2 != 0)
                        {
                            Excep = "X";
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Excep, "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                        if (Excepcion2 == true & Monto2 == 0) {

                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "8", "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                    }
                    if (Excepcion == false)
                    {
                        if (Excepcion2 == false)
                        {
                            matrizpago.viaspago(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "9", "D", "", Convert.ToString(lblPais.Content), Protesto, Condiciones);
                        }
                       
                        if(Excepcion2 == true & Monto2 != 0)
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
                        //cmbVPMedioPag.ItemsSource = null;
                        //cmbVPMedioPag.Items.Clear();

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
                        //cmbVPMedioPag.Items.Add(matrizpago.ObjDatosViasPago);
                        //   cmbVPMedioPag.ItemsSource = matrizpago.ObjDatosViasPago[i].VIA_PAGO + "-" + matrizpago.ObjDatosViasPago[i].DESCRIPCION;

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

                               // Monto = Monto + Convert.ToDouble(partidaseleccionadas[i].MONTOF);

                              Monto = Monto3 + Monto4;
                                //  Monto = Monto + Convert.ToDouble(partidaseleccionadas[i].MONTOF);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                            }
                        }
                        // string mount = string.Format(Convert.ToString(Monto), "##.###,##");
                        // textBlock4.Text = string.Format(Convert.ToString(Monto), "#####,##");
                        cmbBancoProp.ItemsSource = null;
                        cmbBancoProp.Items.Clear();
                        cmbCuentasBancosProp.ItemsSource = null;
                        cmbCuentasBancosProp.Items.Clear();
                        cmbBanco.ItemsSource = null;
                        cmbBanco.Items.Clear();
                        //
                        if (cmbMoneda.Text == "CLP")
                        {
                            string Valor = Convert.ToString(Monto);
                            if (Valor.Contains("-"))
                            {
                                Valor = "-" + Valor.Replace("-", "");
                            }
                            Valor = Valor.Replace(".", "");
                            Valor = Valor.Replace(",", "");
                            decimal ValorAux = Convert.ToDecimal(Valor);
                            string monedachil = string.Format("{0:0,0}", ValorAux);

                            textBlock4.Text = Convert.ToString(monedachil);
                            txtMontoFP.Text = Convert.ToString(monedachil);
                        }
                        else
                        {
                            string moneda = string.Format("{0:0,0.##}", Monto);
                            textBlock4.Text = Convert.ToString(moneda);
                            txtMontoFP.Text = Convert.ToString(moneda);
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
                    //}
                    //HASTA AQUI MODIFICACION A CONDICIONES DE PAGO DISTINTAS
                    // else
                    // {
                    //    DGPagos.UnselectAll();
                    // }
                }
                GC.Collect();
            }
            catch  (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);   
            }
        }
               
        //BOTON QUE DESPLIEGA EL GRID DE DOCUMENTOS POR PAGAR A PARTIR DE LA BUSQUEDA DE UN RUT O NUMERO DE DOCUMENTO
        private void btnBuscarP_Click(object sender, RoutedEventArgs e)
        {

            GBViasPago.Visibility = Visibility.Collapsed;
            LimpiarViasDePago();
            ListaDocumentosPendientes();
            GC.Collect();
           // btnBuscarPM.IsEnabled = false;
           
        }
        //Boton QUE LISTA LA RECAUDACION DE VEHICULOS
        private void btnBuscarR_Click(object sender, RoutedEventArgs e)
        {
            LimpiarViasDePago();
            listaRecaudacionVehiculo();
            GC.Collect();
        }

        //BOTON QUE UBICA EL ARCHIVO A CARGAR EN PAGOS MASIVOS
        private void button3_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel Old files (*.xls)|*.xls|All files (*.*)|*.*";
            if (Convert.ToBoolean(openFileDialog.ShowDialog()) == true)
                txtArchivo.Text = openFileDialog.InitialDirectory;
            //txtArchivo.Text = File.GetAttributes(openFileDialog.InitialDirectory.FileName);
            // txtArchivo.Text = File.ReadAllText(openFileDialog.FileName);
            txtArchivo.Text = openFileDialog.FileName;
            GC.Collect();
        }
        ////BOTON QUE DESPLIEGA EL GRID DE DOCUMENTOS POR PAGAR DE MODO MASIVO A PARTIR DE LA BUSQUEDA DE UN RUT O NUMERO DE DOCUMENTO
        private void btnBuscarPM_Click(object sender, RoutedEventArgs e)
        {
            GBViasPago.Visibility = Visibility.Collapsed;
            //GBDocsAPagar.Visibility = Visibility.Visible;
            LimpiarViasDePago();
            //Lectura del archivo excel de cargas masivas.
            ListaDocumentosPendientesCargasMasivas(txtArchivo.Text);
            GC.Collect();
        }
        //CLICK DEL BOTON QUE MUESTRA LOS DOCUMENTOS PARA REALIZAR ANTICIPOS  
        private void button5_Click(object sender, RoutedEventArgs e)
        {
                GBViasPago.Visibility = Visibility.Collapsed;
                LimpiarViasDePago();
                ListaDocumentosPendientesAnticipos();
                DPFechActual.Text = datePicker1.Text;
                cheques.Clear();
                GC.Collect();
        }      
        //CLICK DEL BOTON DE LA BARRA DE HERRAMIENTAS INICIO
        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            GBInicio.Visibility = Visibility.Visible;
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
			GBrecauda.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
            btnAutAnul.Visibility = Visibility.Collapsed;
            //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
            //LimpiarCamposInformeRendicion();
            GC.Collect();
        }


        //BOTON PARA LA EMISION DE LAS NOTAS DE CREDITO (NC)
        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            GBInicio.Visibility = Visibility.Collapsed;
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Visible;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBrecauda.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Collapsed;
          //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();
          //LimpiarCamposInformeRendicion();
            btnEmitirNC.IsEnabled = false;
            chkNCTribut.IsChecked = false;
            chkNCTribut.IsEnabled = true;
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
		            // You can use the default case.
		           
	            }
            }
            else
            {
                System.Windows.MessageBox.Show("Ingrese el tipo de documento (forma de pago)");
            }
            GC.Collect();
          }
            catch  (Exception ex)
         {
               Console.WriteLine(ex.Message + ex.StackTrace);
               System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
               LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
               GC.Collect();  
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
                //frm.DialogResult = true;// = this;
                frm.Owner = this;
                frm.Show();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }

        }

        private string RemoveSpecialCharacters(string str)
        {
            return Regex.Replace(str, "[^a-zA-Z0-9_.- ]+", "", RegexOptions.Compiled);
        }
        
        
        //BOTON QUE REALIZA EL PAGO Y CREACION DEL COMPROBANTE DE INGRESO DE NOTAS DE VENTAS Y PAGO DE ANTICIPOS
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
                    //DocsAPagar[i].MONTO = DocsAPagar[i].MONTO.Replace(",", ".");
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
                    //pagodocumentosingreso.pagodocumentosingreso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), Convert.ToString(lblSociedad.Content), DocsAPagar[i].NDOCTO, ViasPagoTransaccion, DocsAPagar, Convert.ToString(lblPais.Content), cmbMoneda.Text,Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content));
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
                    //System.Windows.MessageBox.Show(Mensaje);
                    if (pagodocumentosingreso.message != "")
                    {
                        System.Windows.MessageBox.Show(pagodocumentosingreso.message);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show(pagodocumentosingreso.pagomessage);
                    }
                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);

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
                    // for (int i = 0; i < pagoanticipos.T_Retorno.Count; i++)
                    //{
                    //    Mensaje = Mensaje + " - " + pagoanticipos.T_Retorno[i].MESSAGE + " - " + pagodocumentosingreso.T_Retorno[i].MESSAGE_V1;
                    //}
                    if (pagoanticipos.message != "")
                    {
                        System.Windows.MessageBox.Show(pagoanticipos.message);
                    }
                    else
                    {
                        System.Windows.MessageBox.Show(pagoanticipos.status);
                    }
                   
                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);
                    if (pagoanticipos.comprobante != "")
                    {
                        ImpresionesDeDocumentosAutomaticas(pagoanticipos.comprobante,"X");
                        //pagoanticipos.T_Retorno.Clear();
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("No se generó comprobante de pago");
                    }

                    pagoanticipos.T_Retorno.Clear();
                }
                //Limpiar aqui el resumen de las vias de pago
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
          //textBlock4.Text = "";
            textBlock5.Text = "";
            btnConfirPag.IsEnabled = false;
            RFC_Combo_Bancos();
            GC.Collect();
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
                        reimpresionfiscal.ReipresionFiscal2(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprReimp.Text);
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
             
        //BUSQUEDA DE DOCUMENTO POR ANULAR 
        private void btnBuscarAnul_Click(object sender, RoutedEventArgs e)
        {
           txtUserAnula.Text = "";
           txtComentAnula.Text = "";
           LimpiarViasDePago();
           chkFiltro.IsChecked = false;
           if ((txtComprAn.Text == "") && (txtRUTAn.Text == "" ))
           {
               System.Windows.MessageBox.Show("Ingrese un RUT o un número de comprobante");            
           }
           else
           {
               //RFC Y FORM QUE GENERAN VENTANA DE AUTORIZACION SOLO PARA EL SUPERUSUARIO QUE ANULA LOS COMPROBANTES
               ListaDocumentosAnulacion();
           }
           GC.Collect();          
        }

        private void btnBuscarAnulV_Click(object sender, RoutedEventArgs e)
        {
            txtUserAnula.Text = "";
            txtComentAnula.Text = "";
            LimpiarViasDePago();
            chkFiltro.IsChecked = false;
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
        //BOTON QUE EMITE NOTAS DE CREDITO
        public void btnBuscarNC_Click(object sender, RoutedEventArgs e)
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
                    for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
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
        //BOTON QUE LLEVA EL DATO SELECCIONADO DESDE EL MONITOR 
        private void btnPagoMonitor_Click(object sender, RoutedEventArgs e)
        {
            LimpiarViasDePago();
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            ListaDocumentosPendientesDesdeMonitor();
            //btnBuscarP.IsEnabled = false;
            GBDocsAPagar.Visibility = Visibility.Visible;
            GBViasPago.Visibility = Visibility.Visible;
            txtDocu.Text = "";
            txtDocuAnt.Text = "";
            txtRut.Text = "";
            txtRUTAnt.Text = "";
            GC.Collect();
        }
        //ANULACION DE COMPROBANTES
        private void btnAnular_Click(object sender, RoutedEventArgs e)
        {
            //BUSQUEDA DEL COMPROBANTE SELECCIONADO
            string IdComprobante = "";
            //List<CAB_COMP> Comprobante = new List<CAB_COMP>();
            List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
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
            //for (int i = 0; i < DGDocCabec.SelectedItems.Count; i++)
            //{
            //    {
            //        Comprobante.Add(DGDocCabec.SelectedItems[i] as CAB_COMP);
            //    }
            //}


            if (partidaopen.Count == 0)
            {
                System.Windows.MessageBox.Show("Seleccione un comprobante en la tabla de cabeceras");
            }
            else
            {

                //RFC PARA ANULAR COMPROBANTES
                AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, txtUserAnula.Text, txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content));
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
            GC.Collect();
        }

        private void btnAnularV_Click(object sender, RoutedEventArgs e)
        {
            //BUSQUEDA DEL COMPROBANTE SELECCIONADO
            string IdComprobante = "";
            List<CAB_COMP> Comprobante = new List<CAB_COMP>();
            List<CAB_COMPAUX> DocsAPagar = new List<CAB_COMPAUX>();
            DocsAPagar.Clear();
            if (this.DGDocCabec.Items.Count > 0)
            {
                for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
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
            //for (int i = 0; i < DGDocCabec.SelectedItems.Count; i++)
            //{
            //    {
            //        Comprobante.Add(DGDocCabec.SelectedItems[i] as CAB_COMP);
            //    }
            //}
            //IdComprobante = Comprobante[0].ID_COMPROBANTE;

            if (partidaopen.Count == 0)
            {
                System.Windows.MessageBox.Show("Seleccione un comprobante en la tabla de cabeceras");
            }
            else
            {

                //RFC PARA ANULAR COMPROBANTES
                AnulacionComprobantes anulacioncomprobantes = new AnulacionComprobantes();
                anulacioncomprobantes.anulacioncomprobantes(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, IdComprobante, txtUserAnula.Text, txtComentAnula.Text, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content));
                //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
                if (anulacioncomprobantes.Mensaje != "")
                {
                    System.Windows.Forms.MessageBox.Show(anulacioncomprobantes.Mensaje);
                }
                if (anulacioncomprobantes.errormessage != "")
                {
                    System.Windows.Forms.MessageBox.Show(anulacioncomprobantes.errormessage);
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

            }
            GC.Collect();
        }

        #endregion

        //MANEJO DE LOS EVENTOS ASOCIADOS A LOS RADIOBUTTONS
        #region RadioButtons

        
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
            //Limpiar aqui el resumen de las vias de pago
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
            //Limpiar aqui el resumen de las vias de pago
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
            //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
            GC.Collect();
        }

        private void RBDoc_Checked(object sender, RoutedEventArgs e)
        {
            txtRut.Text = "";
            txtRut.Visibility = Visibility.Collapsed;
            txtDocu.Text = "";
            txtDocu.Visibility = Visibility.Visible;
            GBDocsAPagar.Visibility =  Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            partidasabiertas = new PartidasAbiertas();
            btnBuscarP.IsEnabled = true;
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
            GBViasPago.Visibility = Visibility.Visible;
            btnBuscarP.IsEnabled = true;
            //Limpiar aqui el resumen de las vias de pago
            LimpiarViasDePago();
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
            btnBuscarP.IsEnabled = true;
            //Limpiar aqui el resumen de las vias de pago
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
            maestrotarjetas.maestrotarjetas(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content), Convert.ToString(cmbVPMedioPag.Text));
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
                System.Windows.Forms.MessageBox.Show("No existen datos de " + Convert.ToString(cmbVPMedioPag.Text).Substring(3, Convert.ToString(cmbVPMedioPag.Text).Length-3) + " en el sistema");
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
                //cmbBanco.ItemsSource = maestrobancos.T_Retorno[0].BANKL + " - " + maestrobancos.T_Retorno[0].BANKA;
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
                //cmbBanco.ItemsSource = maestrobancos.T_Retorno[0].BANKL + " - " + maestrobancos.T_Retorno[0].BANKA;
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
        //SELECCION DEL COMBOBOX DE BANCOS PROPIOS
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
                                //label40.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                //label40.Margin = new Thickness(146, 33, 0, 0);
                                //txtObserv.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                //txtObserv.Margin = new Thickness(235, 33, 0, 0);
                                //txtObserv.Width = 295;
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
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
            GC.Collect();
        }


        #endregion

        //FUNCIONES y METODOS
        #region // Funciones

        void ImpresionesDeDocumentosAutomaticas(string comprobante, string batch)
        {
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }

            BusquedaReimpresiones busquedareimpresiones = new BusquedaReimpresiones();
            busquedareimpresiones.docsreimpresion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, comprobante, "",LogApert[0].ID_REGISTRO ,Convert.ToString(lblPais.Content),Convert.ToString(textBlock6.Content),"X");

            if (busquedareimpresiones.errormessage != "")
            {
                //System.Windows.Forms.MessageBox.Show(busquedareimpresiones.errormessage);
            }
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
                        , Convert.ToString(textBlock7.Content), Caja, Referencia, Documento, DocContable, InOut, LogApert[0].MONEDA, Pedido, txtMandante.Text);
                if (reimpresioncomprobantes.DatosEmpresa.Count != 0)
                {
                    frm.txtSociedad.Text = reimpresioncomprobantes.DatosEmpresa[0].BUKRS;
                    frm.txtEmpresa.Text = reimpresioncomprobantes.DatosEmpresa[0].BUTXT;
                    frm.txtRIF.Text = reimpresioncomprobantes.DatosEmpresa[0].STCD1;
                }
                frm.Show();
            }
            else
            {
            }
            GC.Collect();
        }
        //FUNCION QUE CONTROLA LA LECTURA DE DATOS DE EL MONITOR
        void timer_Tick(object sender, EventArgs e)
        {
            try
            {
              
                //RFC del Monitor por Timer
                monitor.ObjDatosMonitor.Clear();
                monitor.monitor(Convert.ToString(datePicker1.SelectedDate.Value), Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblSociedad.Content));
                if (monitor.ObjDatosMonitor.Count > 0)
                {
                    DGMonitor.ItemsSource = null;
                    DGMonitor.Items.Clear();
                    DGMonitor.ItemsSource = monitor.ObjDatosMonitor;
                    // MessageBox.Show("Conectandose a la RFC del Monitor por timer c/ 10 seg");
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }
        //FUNCION QUE HACE LA REVISION DEL DIGITO VERIFICADOR
        private string DigitoVerificador(string RUTU)
        {
            string digito= "";
            string RUTsDV = "";
            string RUTcDV = "";
            string RUT = "";
            int Total = 0;
            int j = 2;
            int modulo= 0;
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
           
            for (int i = RUT.Length-1; i >= 0;  --i)
            {
                RUTsDV = RUTsDV + RUT[i];
            }
            for (int i = 0; i < RUTsDV.Length; i++)
            {
                digito = Convert.ToString(RUTsDV[i]);
                Total = Total + Convert.ToInt16(digito)*j;
                if (j < 7)
                {
                    j++;
                }
                else
                {
                    j= 2;
                }
            }
            modulo = Total % 11;
            Total = 11-modulo;
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
                //Eliminación de la separación de miles
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
                
                num_cuotas= Convert.ToInt16(txtCantDoc.Text);
                
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
                        , cta_banco, ifinan,corre, zuonr, hkont, prctr, znop);
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

                                detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago,  Convert.ToInt64(Math.Round(montoaux,0)), moneda, banco, emisor
                                    , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                    , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                    , cta_banco, ifinan,corre, zuonr, hkont, prctr, znop);
                                cheques.Add(detcheq);
                            } 
                        }
                        else //VIAS DE PAGO CON CUOTAS Y MONTO DE LA CUOTAS CON UN RESIDUO QUE SE SUMA EN LA CUOTA FINAL
                        {
                            for (int i = 1; i <= num_cuotas; i++)
                            {
                                double montoaux = Convert.ToDouble(monto / num_cuotas);
                                montotot = montotot +  Math.Round(montoaux,0); 
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
                                //if (MedioPago == "S")
                                    num_cuotasaux = 1;
                                //else
                                //    num_cuotasaux = num_cuotas;
                                if (j != num_cuotas)
                                {
                                    detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, Convert.ToInt64(Math.Round(montoaux, 0)), moneda, banco, emisor
                                     , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                     , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                     , cta_banco, ifinan,corre, zuonr, hkont, prctr, znop);
                                    cheques.Add(detcheq);
                                }
                                else
                                {
                                    detcheq = new DetalleViasPago(mandt, land, id_comprobante, id_detalle, via_pago, Convert.ToInt64(Math.Round(montoaux, 0) + montores), moneda, banco, emisor
                                                                 , num_cheque, cod_autorizacion, Convert.ToString(num_cuotasaux), fecha_venc, texto_posicion, anexo, sucursal, num_cuenta, num_tarjeta
                                                                 , num_vale_vista, patente, num_venta, pagare, fecha_emision, nombre_girador, carta_curse, num_transfer, num_deposito
                                                                 , cta_banco, ifinan,corre, zuonr, hkont, prctr, znop);
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
                                decimal ValorAux = Convert.ToDecimal(MntTotalChq);
                                string monedachil = string.Format("{0:0,0}", ValorAux);
                               items.Add(new MontoMediosdePago(Convert.ToString(cmbVPMedioPag.Items[i]), monedachil));
                            }
                            else
                            {
                                decimal ValorAux = Convert.ToDecimal(MntTotalChq);
                                string monedaforex = string.Format("{0:0,0.##}", ValorAux);
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
                    TotalVPagos = TotalVPagos + Convert.ToDouble(items[i].Monto);
                }
                double MntTotalPend = 0;
                for (int i = 0; i <= partidaseleccionadas.Count - 1; i++)
                {
                    MntTotalPend =  MntTotalPend + Convert.ToDouble(partidaseleccionadas[i].MONTOF_PAGAR);
                }
               
                if (moneda == "CLP")
                {
                    decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = string.Format("{0:0,0}", ValorAux);
                    textBlock3.Text = Convert.ToString(monedachil);
                    decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = string.Format("{0:0,0}", ValorAux2);
                    textBlock4.Text = Convert.ToString(monedachil2);
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = string.Format("{0:0,0}", ValorAux3);
                    if (monedachil3 == "00")
                    {
                        monedachil3 = "0";
                    }
                    textBlock5.Text = Convert.ToString(monedachil3);                
                }
                else
                {
                    decimal ValorAux = Convert.ToDecimal(TotalVPagos);
                    string monedachil = string.Format("{0:0,0.##}", ValorAux);
                    textBlock3.Text = Convert.ToString(monedachil);
                    decimal ValorAux2 = Convert.ToDecimal(MntTotalPend);
                    string monedachil2 = string.Format("{0:0,0.##}", ValorAux2);
                    textBlock4.Text = Convert.ToString(monedachil2);
                    decimal ValorAux3 = Convert.ToDecimal((MntTotalPend) - (TotalVPagos));
                    string monedachil3 = string.Format("{0:0,0.##}", ValorAux3);
                    textBlock5.Text = Convert.ToString(monedachil3);   
                    //PART_ABIERTAS_resp.MONTOF_PAGAR = string.Format("{0:0,0.##}", ValorAux);
                }

                //textBlock3.Text = Convert.ToString(TotalVPagos);
                //textBlock4.Text = Convert.ToString(MntTotalPend);
                //textBlock5.Text = Convert.ToString((MntTotalPend) - (TotalVPagos));
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
                        if (textBlock5.Text == "0")
                            {
                                btnConfirPag.IsEnabled = true;
                            }
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
                    //cmbVPMedioPag.Text = "";
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
                    //txtCodAuto.Text = "";
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
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
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
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
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
                DatPckPgDoc.Text = datePicker1.Text;
            }
            //Se busca el dato del documento en la grilla del monitor
            try
            {
               
                //Calculo del monto para los documentos y partidas abiertas seleccionadas.
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

                ////***RFC Partidas abiertas para pago busqueda por RUT
                //Calculo del monto para los documentos y partidas abiertas seleccionadas.
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
                                monitorseleccionado[i].MONTOF =monitorseleccionado[i].MONTOF.Substring(posicion, 1) + monitorseleccionado[i].MONTOF.Substring(0, posicion);
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
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                
            }
            if (chkMonitor.IsChecked.Value)
            {
               // timer.Start();
            }
            GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect(); 
            }

        }


        //FUNCION QUE TRAE LOS DOCUMENTOS PENDIENTES O PARTIDAS ABIERTAS POR CARGAS MASIVAS
        private void ListaDocumentosPendientesCargasMasivas(string thisFileName)
        {
              try
            {
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
                string RutExc = "";
                string SocExc = "";
                List<PagosMasivosNuevo> ListaExc = new List<PagosMasivosNuevo>();
                
                //LLAMADO A LA FUNCION QUE LEE EL ARCHIVO EXCEL
                RecogerDatosExcel2(thisFileName, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ref SocExc, ref RutExc, out ListaExc, ref  PrgBarExcel);
              

               //RecogerDatosExcel(thisFileName, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content),  ref SocExc, ref RutExc, out ListaExc,ref  PrgBarExcel);
              //RFC que hace la compensacion de pagos masivos
                PagosMasivosNew pagosmasivos = new PagosMasivosNew();
                pagosmasivos.pagosmasivos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                    , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(lblPais.Content)
                    , Convert.ToString(DateTime.Today), thisFileName, LogApert[0].ID_REGISTRO, LogApert[0].ID_CAJA, LogApert[0].MONEDA, ListaExc); 
               ////RFC que trae los documentos leidos del archivo excel de cargas masivas
               // documentospagosmasivos.pagosmasivos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, RutExc, SocExc, ListaExc);

               // DGPagos.ItemsSource = null;
               // DGPagos.Items.Clear();
               // if (documentospagosmasivos.ObjDatosPartidasOpen.Count > 0)
               // {
               //    DGPagos.ItemsSource = documentospagosmasivos.ObjDatosPartidasOpen;x
               // }
                if (pagosmasivos.message != "")
                {
                    System.Windows.Forms.MessageBox.Show(pagosmasivos.message);
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
                  LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                  GC.Collect();
              }
        }


        //FUNCION QUE TRAE LOS DOCUMENTOS A PAGAR A PARTIR DE UN ARCHIVO EXCEL
        static void RecogerDatosExcel2(string ruta, string usuario, string sucursal, string idcaja, ref string SocExc, ref string RutExc, out List<PagosMasivosNuevo> ListaExc, ref System.Windows.Controls.ProgressBar PrgBarExcel)
        {
            ListaExc = new List<PagosMasivosNuevo>();

            //Declaro las variables necesarias/
            Microsoft.Office.Interop.Excel._Application xlApp;
            Microsoft.Office.Interop.Excel._Workbook xlLibro;
            Microsoft.Office.Interop.Excel._Worksheet xlHoja1;
            Microsoft.Office.Interop.Excel.Sheets xlHojas;
            // Microsoft.Office.Interop.Excel.Sheets xlHojas2;


            //asigno la ruta dónde se encuentra el archivo
            string fileName = ruta;
            // inicializo la variable xlApp (referente a la aplicación)
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            //Muestra la aplicación Excel si está en true
            xlApp.Visible = false;
            // Abrimos el libro a leer (documento excel)
            xlLibro = xlApp.Workbooks.Open(fileName);
            try
            {
                //Asignamos las hojas
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
                    //recorremos las celdas que queremos y sacamos los datos 
                    //10 es el número de filas que queremos que lea
                    // System.Windows.Controls.ProgressBar  PrgBarExcel = new System.Windows.Controls.ProgressBar();
                    PrgBarExcel.Value = 0;
                    string row = "";// (string)xlHoja1.Cells[j, "A"].Text;
                    string col = "";//(string)xlHoja1.Cells[j, "B"].Text;
                    string value = "";// (string)xlHoja1.Cells[j, "B"].Text;

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
                    //        TextInput("Cargando registros...");
                    for (int i = 3; i <= n; i++)
                    {
                      if (verificador >= 2)
                        {
                            break;
                        }
                         
                        if (((string)xlHoja1.Cells[j, "A"].Text != "") && ((string)xlHoja1.Cells[j, "B"].Text != ""))
                        {
                            

                            //pagosm = new PagosMasivosNuevo(row, col, value);
                            for (int r= 1; r<=2; r++)
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
                    
                    value = (string)xlHoja1.Cells[Convert.ToString(xlHoja1.UsedRange.Rows.Count), "B" ].Text;
                    pagosm.COL = col;
                    pagosm.ROW = row;
                    pagosm.VALUE = value;
                    ListaExc.Add(pagosm);
                }


                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, usuario, idcaja, sucursal, ex.Message + ex.StackTrace);

                }
                if (ListaExc.Count == 0)
                {
                    //return ListaExc;
                    System.Windows.MessageBox.Show("Error en el archivo excel a cargar. Revise formato del archivo o el formato de la plantilla o si el archivo tiene datos");
                }
                else
                {

                    // return ListaExc;
                    System.Windows.MessageBox.Show(Convert.ToString(ListaExc.Count) + " documentos cargados");
                }
            }

            finally
            {
                //Cerrar el Libro
                xlLibro.Close(false);
                //Cerrar la Aplicación
                xlApp.Quit();
                PrgBarExcel.Value = 0;
                GC.Collect();
            }
        }

        static void RecogerDatosExcel(string ruta, string usuario, string sucursal, string idcaja, ref string SocExc, ref string RutExc, out List<PagosMasivos> ListaExc, ref System.Windows.Controls.ProgressBar PrgBarExcel) 
        {
            ListaExc = new List<PagosMasivos>();

            //Declaro las variables necesarias/
            Microsoft.Office.Interop.Excel._Application xlApp;
            Microsoft.Office.Interop.Excel._Workbook xlLibro;
            Microsoft.Office.Interop.Excel._Worksheet xlHoja1;
            Microsoft.Office.Interop.Excel.Sheets xlHojas;
           // Microsoft.Office.Interop.Excel.Sheets xlHojas2;
            
           
            //asigno la ruta dónde se encuentra el archivo
            string fileName = ruta;
           // inicializo la variable xlApp (referente a la aplicación)
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            //Muestra la aplicación Excel si está en true
            xlApp.Visible = false;
           // Abrimos el libro a leer (documento excel)
            xlLibro = xlApp.Workbooks.Open(fileName);
            try
            {
                //Asignamos las hojas
                xlHojas = xlLibro.Sheets;
             
                try
                {
                    int k = 1;
                    //Asignamos la hoja con la que queremos trabajar: 
                    xlHoja1 = (Microsoft.Office.Interop.Excel._Worksheet)xlHojas["Hoja1"];
                    int n = xlHoja1.UsedRange.Rows.Count;
                    PrgBarExcel.Maximum = n;
                    int j = 4;
                    int m = 2;
                    int l = 1;
                    int verificador = 0;
                   SocExc = "";
                   RutExc = "";
                    //recorremos las celdas que queremos y sacamos los datos 
                    //10 es el número de filas que queremos que lea
                // System.Windows.Controls.ProgressBar  PrgBarExcel = new System.Windows.Controls.ProgressBar();
                   PrgBarExcel.Value = 0;
                   
                   //        TextInput("Cargando registros...");
                    for (int i = 3; i <= n; i++)
                    {
                        //string referencia = (string)xlHoja1.get_Range("A"+j).Text;
                        //string monto = (string)xlHoja1.get_Range("B"+j).Text;
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
                            //System.Threading.Thread.Sleep(100);
                           
                            
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
                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, usuario, idcaja, sucursal, ex.Message + ex.StackTrace);

                }
                if (ListaExc.Count == 0)
                {
                   //return ListaExc;
                    System.Windows.MessageBox.Show("Error en el archivo excel a cargar. Revise formato del archivo o el formato de la plantilla o si el archivo tiene datos");
                }
                else
                {

                   // return ListaExc;
                System.Windows.MessageBox.Show(Convert.ToString(ListaExc.Count) + " documentos cargados");
                }
            }

            finally
            {
                //Cerrar el Libro
                xlLibro.Close(false);
                //Cerrar la Aplicación
                xlApp.Quit();
                PrgBarExcel.Value = 0;
                GC.Collect();
            }
   
            
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message + ex.StackTrace);
            //}
        }

        //FUNCION QUE TRAE LOS DOCUMENTOS PARA NOTAS DE CREDITO A PARTIR DE LA BUSQUEDA POR RUT O DOCUMENTO
        private void ListaDocumentosNC()
        {
            try
            {
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }

                partidasabiertas.ObjDatosPartidasOpen.Clear();
                DGPagos.ItemsSource = null;
                DGPagos.Items.Clear();

                // textBlock6.Content = cajaconect;
                // textBlock7.Content = usuariologg;

                if (DatPckPgDoc.Text == "")
                {
                    DatPckPgDoc.Text = datePicker1.Text;
                }

                // PartidasAbiertas partidasabiertas = new PartidasAbiertas();

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
                    Documento = txtDocu.Text;
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
                        txtDocu.Text = Documento;

                        notasdecredito.notasdecredito(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprNC.Text, txtRUTNC.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "Documento", Convert.ToString(textBlock6.Content));
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }

                if (notasdecredito.ObjDatosNC.Count > 0)
                {
                    //GBDocsAPagar.Visibility = Visibility.Visible;
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
                    //DGDocCabec.ItemsSource = notasdecredito.ObjDatosNC;
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
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }

        }

        //FUNCION QUE TRAE LOS DOCUMENTOS A PAGAR A PARTIR DE LA BUSQUEDA POR RUT O DOCUMENTO
        private void ListaDocumentosPendientes()
        {
            try
            {
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }

                partidasabiertas.ObjDatosPartidasOpen.Clear();
                DGPagos.ItemsSource = null;
                DGPagos.Items.Clear();

                // textBlock6.Content = cajaconect;
                // textBlock7.Content = usuariologg;

                if (DatPckPgDoc.Text == "")
                {
                    DatPckPgDoc.Text = datePicker1.Text;
                }

                // PartidasAbiertas partidasabiertas = new PartidasAbiertas();
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

                        partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "", txtDocu.Text, txtRut.Text, Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text), Convert.ToString(lblPais.Content), "", "RUT");
                    }
                }
                else if (RBDoc.IsChecked == true)
                {
                    //***RFC Partidas abiertas para pago busqueda por numero de documento
                    string Documento = "";
                    Documento = txtDocu.Text;
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
                        txtDocu.Text = Documento;

                        partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, "", txtDocu.Text, txtRut.Text, Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text), Convert.ToString(lblPais.Content), "", "Documento");
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }

                if (partidasabiertas.ObjDatosPartidasOpen.Count > 0)
                {
                    GBDocsAPagar.Visibility = Visibility.Visible;
                    DGPagos.ItemsSource = null;
                    DGPagos.Items.Clear();
                    List<T_DOCUMENTOSAUX> partidaopen = new List<T_DOCUMENTOSAUX>();

                    for (int k = 0; k < partidasabiertas.ObjDatosPartidasOpen.Count; k++)
                    {
                        T_DOCUMENTOSAUX partOpen = new T_DOCUMENTOSAUX();
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
                    }
                   


                   // DGPagos.ItemsSource = partidasabiertas.ObjDatosPartidasOpen;
                        DGPagos.ItemsSource = partidaopen;
                }
                GC.Collect();
            }
            catch  (Exception ex)
            {
               Console.WriteLine(ex.Message + ex.StackTrace);
               System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
               LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
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
           
           // String RUT = "11601174-3";
            if (RBRutRE.IsChecked == true)
            {
                if (RUT != RUTAux)
                {
                    System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                }
                else
                {
                    recauda.recauVehi(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocu.Text, RUT, Convert.ToString(lblSociedad.Content));

                    bapi_return2 = recauda.objReturn2;

                    for (int i = 0; i < bapi_return2.Count(); i++)
                    {
                        
                        switch (bapi_return2[i].TYPE)
                        {
                            case "E":
                                //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_stop.gif' />&nbsp;";
                                break;
                            case "I":
                                //cadImagen = "&nbsp;<img src='../../../Images/info.gif' />&nbsp;";
                                break;
                            case "W":
                                //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_warning.gif' />&nbsp;";
                                break;
                            case "S":
                                //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_success.gif' />&nbsp;";
                                break;
                        }
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
                txtDocu.Text = Documento;

                recauda.recauVehi(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocu.Text, RUT, Convert.ToString(lblSociedad.Content));
                bapi_return2 = recauda.objReturn2;

                for (int i = 0; i < bapi_return2.Count(); i++)
                {
                    switch (bapi_return2[i].TYPE)
                    {
                        case "E":
                            //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_stop.gif' />&nbsp;";
                            break;
                        case "I":
                            //cadImagen = "&nbsp;<img src='../../../Images/info.gif' />&nbsp;";
                            break;
                        case "W":
                            //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_warning.gif' />&nbsp;";
                            break;
                        case "S":
                            //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_success.gif' />&nbsp;";
                            break;
                    }
                    cadMensajes2 = cadMensajes2 + bapi_return2[i].MESSAGE + "<br>";
                    System.Windows.MessageBox.Show(cadMensajes2);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
            }

            if(recauda.objPag.Count > 0)
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
                //DGRecau.ItemsSource = recauda.objPag;
                DGAutorizacionVehiculos.ItemsSource = null;
                DGAutorizacionVehiculos.Items.Clear();

            }
            GC.Collect();
        }

		//FUNCION QUE TRAE LOS DOCUMENTOS A PAGAR POR ANTICIPOS A PARTIR DE LA BUSQUEDA POR RUT O DOCUMENTO
        private void ListaDocumentosPendientesAnticipos()
        {
            try
            {
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }

                anticipos.ObjDatosAnticipos.Clear();
                DGPagos.ItemsSource = null;
                DGPagos.Items.Clear();

                // textBlock6.Content = cajaconect;
                // textBlock7.Content = usuariologg;

                if (DatPckPgDoc.Text == "")
                {
                    DatPckPgDoc.Text = datePicker1.Text;
                }

                // PartidasAbiertas partidasabiertas = new PartidasAbiertas();
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
                    //Verificacion de RUT
                    if (RUT != RUTAux)
                    {
                        System.Windows.Forms.MessageBox.Show("Número de RUT inválido");
                        txtRUTAnt.Focus();
                    }
                    else
                    {

                        anticipos.anticiposopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocuAnt.Text, txtRUTAnt.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "RUT");
                       // partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), "", txtDocu.Text, txtRut.Text, Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text), Convert.ToString(lblPais.Content), "", "RUT");
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

                        //***RFC Anticipos para pago busqueda por numero de documento
                        anticipos.anticiposopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtDocuAnt.Text, txtRUTAnt.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), "Documento");
                    }
                    //partidasabiertas.partidasopen(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), "", txtDocu.Text, txtRut.Text, Convert.ToString(lblSociedad.Content), Convert.ToDateTime(DatPckPgDoc.Text), Convert.ToString(lblPais.Content), "", "Documento");
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


                   //DGPagos.ItemsSource = anticipos.ObjDatosAnticipos;
                    DGPagos.ItemsSource = partidaopen;
                   

                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }

        //FUNCION QUE TRAE LOS DOCUMENTOS A ANULAR
        public void ListaDocumentosAnulacion()
        {
            try
            {
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }
                BusquedaAnulacion busquedaanulacion = new BusquedaAnulacion();
               
               
                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();

                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();

                btnRevisDoc.Visibility = Visibility.Collapsed;

                // textBlock6.Content = cajaconect;
                // textBlock7.Content = usuariologg;

              
                // PartidasAbiertas partidasabiertas = new PartidasAbiertas();
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
                    //***RFC Partidas abiertas para ANULACION busqueda por RUT
                    String RUT = DigitoVerificador(RUTAux);
                    //Verificacion de RUT
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
                    //***RFC Partidas abiertas para ANULACION busqueda por numero de documento
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

                        busquedaanulacion.docsanulacion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprAn.Text, txtRUTAn.Text, Convert.ToString(lblSociedad.Content), Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "A");
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }
                //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
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

                        partidaopen.Add(partOpen);
                    }

                    DGDocCabec.ItemsSource = partidaopen;
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = busquedaanulacion.DetalleDocs;
                    btnAnular.IsEnabled = true;
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }

        public void ListaDocumentosAnulacionVehiculos()
        {
            try
            {
                //for (int i = 0; i <= detalledocs.Count - 1; i++)
                for (int i = detalledocs.Count - 1; i >= 0; --i)
                {
                    detalledocs.RemoveAt(i);
                }
                BusquedaAnulacion busquedaanulacion = new BusquedaAnulacion();


                DGDocCabec.ItemsSource = null;
                DGDocCabec.Items.Clear();

                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();
                btnRevisDoc.Visibility = Visibility.Collapsed;

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
                    //***RFC Partidas abiertas para ANULACION busqueda por RUT
                    String RUT = DigitoVerificador(RUTAux);
                    //Verificacion de RUT
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
                    //***RFC Partidas abiertas para ANULACION busqueda por numero de documento
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
                //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
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
                        
                        partidaopen.Add(partOpen);
                    }

                    DGDocCabec.ItemsSource = partidaopen;
                    //DGDocCabec.ItemsSource = busquedaanulacion.CabeceraDocs;
                    DGDocDet.ItemsSource = null;
                    DGDocDet.Items.Clear();
                    DGDocDet.ItemsSource = busquedaanulacion.DetalleDocs;
                    btnAnularV.IsEnabled = true;
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
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

                // textBlock6.Content = cajaconect;
                // textBlock7.Content = usuariologg;
                List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
                for (int i = 1; i < DGLogApertura.Items.Count; i++)
                {
                    if (i == 1)
                    {
                        DGLogApertura.Items.MoveCurrentToFirst();
                    }
                    if (DGLogApertura.Items.CurrentItem != null)
                    {
                        LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                    }
                    DGLogApertura.Items.MoveCurrentToNext();
                }

                // PartidasAbiertas partidasabiertas = new PartidasAbiertas();
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

                  
                    //Verificacion de RU

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

                        busquedareimpresiones.docsreimpresion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprReimp.Text, txtRUTReimp.Text,LogApert[0].ID_REGISTRO ,Convert.ToString(lblPais.Content),Convert.ToString(textBlock6.Content),"");
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

                        busquedareimpresiones.docsreimpresion(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtComprReimp.Text, txtRUTReimp.Text, LogApert[0].ID_REGISTRO, Convert.ToString(lblPais.Content), Convert.ToString(textBlock6.Content), "");
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show("Seleccione una forma de búsqueda por RUT o Número de documento");
                }
                //Limpieza de los datagrid con la data de cabecera y detalle de los comprobantes a anular 
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
                   //DGDocCabec.ItemsSource = busquedareimpresiones.Documentos;
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
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), ex.Message + ex.StackTrace);
                GC.Collect();
            }
        }


        private void limpiar()
        {
            DGRecau.ItemsSource = null;
            DGRecau.Items.Clear();
            DGAutorizacionVehiculos.ItemsSource = null;
            DGAutorizacionVehiculos.Items.Clear();
            GC.Collect();
        }

        private void LimpiarElementosDeCierreDeCaja()
        {
            //Limpiar aqui el resumen de las vias de pago
            txtCommCierre.Text = "";
            txtCommDif.Text = "";
            txtC1.Text = "0";
            txtC10.Text = "0";
            txtC100.Text = "0";
            txtC1000.Text = "0";
            txtC10000.Text = "0";
            txtC5.Text = "0";
            txtC50.Text = "0";
            txtC500.Text = "0";
            txtC5000.Text = "0";
            txtC2000.Text = "0";
            txtC20000.Text = "0";
            DGResumenCaja.ItemsSource = null;
            DGResumenCaja.Items.Clear();
            txtTotalCaja.Text = "";
            txtTotalEfectivo.Text = "";
            txtDiferencia.Text = "";
            txtTotEfect.Text = "";
            txtMApp.Text = "";
            txtMChqDia.Text = "";
            txtMChqFech.Text = "";
            txtMCredit.Text = "";
            txtMDepos.Text = "";
            txtMEfect.Text = "";
            txtMEgresos.Text = "";
            //txtMFFijo.Text = "";
            txtMFinanc.Text = "";
            txtMIngresos.Text = "";
            txtMSaldoF.Text = "";
            txtMTarj.Text = "";
            txtMTransf.Text = "";
            txtMValeV.Text = "";
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
            //txtDocu.Text = "";
            //txtDocuAnt.Text = "";
            //txtRut.Text = "";
            //txtRUTAnt.Text = "";
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
            timer.Stop();
            List<LOG_APERTURA> logapertura = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    logapertura.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }

                      
            BloquearCaja bloquearcaja = new BloquearCaja();
            bloquearcaja.bloqueardesbloquearcaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, logapertura);

            this.Close();
            GC.Collect();
        }

        private void chkDocFiscales_Checked(object sender, RoutedEventArgs e)
        {
            RBDocReimp.IsChecked = true;
            txtComprReimp.Text = "";
            txtRUTReimp.Text = "";
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
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

        private void chkDocFiscales_Unchecked(object sender, RoutedEventArgs e)
        {
            RBDocReimp.IsChecked = false;
            txtComprReimp.Text = "";
            txtRUTReimp.Text = "";
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            DGPagos.ItemsSource = null;
            DGPagos.Items.Clear();
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocCabec.Items.Clear();
            //btnReimpr.IsEnabled = true;           
            btnReimpr.Visibility = Visibility.Visible;
            btnBuscarReimp.Content = "Buscar";
            GC.Collect();
        }

        private void btnRendir_Click(object sender, RoutedEventArgs e)
        {
                       
           List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
           for (int i = 1; i < DGLogApertura.Items.Count; i++)
           {
              if (i == 1)
              {
                  DGLogApertura.Items.MoveCurrentToFirst();
              }
              if (DGLogApertura.Items.CurrentItem != null)
              {
                  LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
              }
              DGLogApertura.Items.MoveCurrentToNext();
           }

            
            //RFC Rendicion Caja
            RendicionCaja rendicioncaja = new RendicionCaja();
            rendicioncaja.rendicioncaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                , DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content)
                , Convert.ToString(lblSociedad.Content), LogApert[0].ID_REGISTRO, "0000000000", "0000000000", LogApert[0].MONEDA);

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
                GBCierreCaja.Visibility = Visibility.Visible;
                GBCommentCierre.Visibility = Visibility.Visible;
                DGResumenCaja.ItemsSource = null;
                DGResumenCaja.Items.Clear();
                DGResumenCaja.ItemsSource = rendicioncaja.detalle_rend;

                if (cmbMoneda.Text == "CLP")
                {
                    string Valor = Convert.ToString(rendicioncaja.MontoEfect);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    decimal ValorAux = Convert.ToDecimal(Valor);
                    string monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMEfect.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoChqDia);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMChqDia.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoChqFech);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMChqFech.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoTransf);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMTransf.Text = Convert.ToString(monedachil);
                    Valor = Convert.ToString(rendicioncaja.MontoValeV);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMValeV.Text = Convert.ToString(monedachil);
                    Valor = Convert.ToString(rendicioncaja.MontoDepot);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMDepos.Text = Convert.ToString(monedachil);
                    Valor = Convert.ToString(rendicioncaja.MontoTarj);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    //
                    txtMTarj.Text = Convert.ToString(monedachil);
                    Valor = Convert.ToString(rendicioncaja.MontoFinanc);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMFinanc.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoApp);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMApp.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoCredit);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMCredit.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoEgresos);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMEgresos.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoIngresos);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMIngresos.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.MontoCCurse);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMCartaC.Text = Convert.ToString(monedachil);
                    //
                    Valor = Convert.ToString(rendicioncaja.SaldoTotal);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtMSaldoF.Text = monedachil;
                    //
                    Valor = Convert.ToString(rendicioncaja.SaldoTotal);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    ValorAux = Convert.ToDecimal(Valor);
                    monedachil = string.Format("{0:0,0}", ValorAux);
                    txtTotalCaja.Text = monedachil;
                }
                else
                {
                    string moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoEfect));
                    txtMEfect.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoChqDia));
                    txtMChqDia.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoChqFech));
                    txtMChqFech.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoTransf));
                    txtMTransf.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoValeV));
                    txtMValeV.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoDepot));
                    txtMDepos.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoTarj));
                    txtMTarj.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoFinanc));
                    txtMFinanc.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoApp));
                    txtMApp.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoCredit));
                    txtMCredit.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoEgresos));
                    txtMEgresos.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoIngresos));
                    txtMIngresos.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.MontoCCurse));
                    txtMCartaC.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.SaldoTotal));
                    txtMSaldoF.Text = moneda;
                    moneda = string.Format("{0:0,0.##}", Convert.ToString(rendicioncaja.SaldoTotal));
                    txtTotalCaja.Text = moneda;
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
                //btnPreCierre.IsEnabled = true;
                btnInformePreCierre.IsEnabled = true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos para el informe de rendición");
            }
            if ((rendicioncaja.id_arqueo != "0000000000") & (rendicioncaja.id_cierre != "0000000000"))
            {
                //btnGestionDep.IsEnabled = false;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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
                DetArqueo.FECHA_REND = Convert.ToString(datePicker1.SelectedDate);
                DetArqueo.MONEDA = cmbMoneda.Text;
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

        private void btnCalcArqueo_Click(object sender, RoutedEventArgs e)
        {
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
            List<DETALLE_ARQUEO> DetalleEfectivo = new List<DETALLE_ARQUEO>(); 
            DetalleEfectivo = ListaEfectivo();

            ArqueoCaja arqueoCaja = new ArqueoCaja();
            arqueoCaja.arqueocaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), cmbMoneda.Text, LogApert[0].ID_REGISTRO, "0000000000", "", "0000000000", LogApert[0].MONTO, DetalleEfectivo);

            if (arqueoCaja.id_arqueo != "0000000000")
            {
                //btnCerrarCaja.IsEnabled = true;
                //txtArqueo.Text = arqueoCaja.id_arqueo;
            }
           
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
            if (cmbMoneda.Text == "CLP")
            {
                    string Valor = Convert.ToString(arqueoCaja.diferencia);
                    if (Valor.Contains("-"))
                    {
                        Valor = "-" + Valor.Replace("-", "");
                    }
                    Valor = Valor.Replace(".", "");
                    Valor = Valor.Replace(",", "");
                    decimal ValorAux = Convert.ToDecimal(Valor);
                    string monedachil = string.Format("{0:0,0}", ValorAux);
                    txtDiferencia.Text = monedachil;
            }
            else
            {
                string moneda = string.Format("{0:0,0.##}", Convert.ToString(arqueoCaja.diferencia));
                txtDiferencia.Text = moneda;
              //txtDiferencia.Text = arqueoCaja.diferencia;
            }

            GC.Collect();  
        }

        private void btnArqueo_Click(object sender, RoutedEventArgs e)
        {
           List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
           for (int i = 1; i < DGLogApertura.Items.Count; i++)
           {
              if (i == 1)
              {
                  DGLogApertura.Items.MoveCurrentToFirst();
              }
              if (DGLogApertura.Items.CurrentItem != null)
              {
                  LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
              }
              DGLogApertura.Items.MoveCurrentToNext();
           }
          
          List<DETALLE_ARQUEO> DetalleEfectivo = new List<DETALLE_ARQUEO>();


           DetalleEfectivo = ListaEfectivo();
            ArqueoCaja arqueoCaja = new ArqueoCaja();
            arqueoCaja.arqueocaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), cmbMoneda.Text, LogApert[0].ID_REGISTRO, "0000000000", "A", "0000000000", LogApert[0].MONTO, DetalleEfectivo);

            if (arqueoCaja.message != "")
            {
              //  System.Windows.MessageBox.Show(arqueoCaja.message);
            }
            if (arqueoCaja.errormessage !="")
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
               // txtDiferencia.Text = arqueoCaja.diferencia;
            }
            FormatoMonedas FM = new FormatoMonedas();
             if (cmbMoneda.Text == "CLP")
            { 
                string Formateo = FM.FormatoMonedaChilena(Convert.ToString(arqueoCaja.diferencia),"1");
                txtDiferencia.Text = Formateo;
            }
            else
            {
                //string moneda = string.Format("{0:0,0.##}", Convert.ToString(arqueoCaja.diferencia));
                string Formateo = FM.FormatoMonedaExtranjera(Convert.ToString(arqueoCaja.diferencia));
                txtDiferencia.Text = Formateo;
            }
            //txtDiferencia.Text = arqueoCaja.diferencia;
            GC.Collect();
        }

        private void ImpresionInformePreCierre()
        {
            try
            {
                Document pdfcommande = new Document(PageSize.LETTER.Rotate(), 20f, 20f, 20f, 20f);;
                string direct = Convert.ToString(System.IO.Path.GetTempPath());
                string fecha = Convert.ToString(DateTime.Today);
                fecha = fecha.Replace("/", "");
                fecha = fecha.Substring(0, 8);
                direct = direct + "InduLog\\" + "PrecierreCaja" + fecha + ".pdf";
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
                                PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NUM_CHEQUE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow7);
                                PdfPCell cellrow8 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_DIA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow8);
                                PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_FECHA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow9);
                                PdfPCell cellrow10 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_TRANSF), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow10);
                                PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_VALE_V), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow11);
                                PdfPCell cellrow12 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_DEP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow12);
                                PdfPCell cellrow13 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_TARJ), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow13);
                                PdfPCell cellrow14 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_FINANC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow14);
                                PdfPCell cellrow15 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_APP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow15);
                                PdfPCell cellrow16 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_CREDITO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                cellrow6.Left = 2f;
                                table2.AddCell(cellrow16);
                                PdfPCell cellrow17 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].PATENTE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow17);
                                PdfPCell cellrow18 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONTO_C_CURSE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                cellrow8.Left = 5f;
                                //cellrow8.
                                table2.AddCell(cellrow18);
                                PdfPCell cellrow19 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DOC_SAP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                table2.AddCell(cellrow19);
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
                    //RendicionCaja rendi = new RendicionCaja();
                    //InfoSociedad = rendi.ObjInfoSoc;

                    //for (int l = 0; l < InfoSociedad.Count(); l++)
                    //{
                    //    texto = InfoSociedad[l].BUKRS;
                    //    texto1 = InfoSociedad[l].STCD1;
                    //}
                    //Datos de empresa
                    
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

                   
                    //pdfcommande.Add(table);
                    //pdfcommande.Add(new iTextSharp.text.Paragraph(Convert.ToString(label11.Content)));
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
                //string url_reimpresion = "";
                //url_reimpresion = direct;
                //PDFViewer pdfvisor = new PDFViewer();
                //pdfvisor.webBrowser1.Navigate(url_reimpresion);
                //pdfvisor.Owner = this;
                //pdfvisor.Show();
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
            }

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
        //CIERRE DE CAJA PROVISIONAL  (BOTON QUEDARA OCULTO PARA LAS VERSIONES DEFINITIVAS POSTERIORES
        private void btnCerrarCaja_Click(object sender, RoutedEventArgs e)
        {
           

            //***RFC cierre de Caja
            CierreCaja cierrecaja = new CierreCaja();
            cierrecaja.cierrecaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(lblPais.Content), "5000", "1000", "Probando 1", "Probando 2");
            System.Windows.Forms.MessageBox.Show(cierrecaja.T_Retorno[0].MESSAGE.ToString());
            if (cierrecaja.status == "S")
            {
                this.IsEnabled = false;
                this.Close();

               // MainWindow frm = new MainWindow();
                //frm.Visibility = Visibility.Visible;
                //frm.Show();

            }
            GC.Collect();
        }
        //CIERRE DE CAJA DEFINITIVO
        private void btnCierreCaja_Click(object sender, RoutedEventArgs e)
        {
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
            //***RFC cierre de Caja
            CierreCajaDefinitivo cierrecaja = new CierreCajaDefinitivo();
            cierrecaja.cierrecajadefinitivo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content),LogApert[0].ID_REGISTRO, txtTotalCaja.Text, txtDiferencia.Text, txtCommDif.Text, txtCommCierre.Text, LogApert[0].MONTO
                , "C", txtArqueo.Text);

            if (cierrecaja.errormessage != "")
            {
                System.Windows.Forms.MessageBox.Show( cierrecaja.errormessage);
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
                LimpiarElementosDeCierreDeCaja();
            }

            //Si existe ID de Cierre, bloquear botones de Menu salvo gestion de depositos

            GC.Collect();
        }


        private void LimpiarCamposInformeRendicion()
        {

            DGResumenCaja.ItemsSource = null;
            DGResumenCaja.Items.Clear();


            txtMEfect.Text = "";
            txtMChqDia.Text = "";
            txtMChqFech.Text = "";
            txtMTransf.Text = "";
            txtMValeV.Text = "";
            txtMDepos.Text = "";
            txtMTarj.Text = "";
            txtMFinanc.Text = "";
            txtMApp.Text = "";
            txtMCredit.Text = "";
            txtMEgresos.Text = "";
            txtMIngresos.Text = "";
            //txtMFFijo.Text = "";
            txtMSaldoF.Text = "";
            GC.Collect();
        }

  


		private void btnActu_Click(object sender, RoutedEventArgs e)
        {

            listaRecaudacionVehiculo();
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
            List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> partidaopen = new List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();

            for (int k = 0; k < DocsAPagar.Count; k++)
            {
                if (DocsAPagar[k].ISSELECTED == true)
                {
                    CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS partOpen = new CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS();
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


            if (NCCheck.Efectivo != "")
            {
                txtTotEfect.Text = NCCheck.Efectivo;
            }
            else
            {
                txtTotEfect.Text = "0";
            }

            if (NCCheck.errormessage != "")
            {
                System.Windows.Forms.MessageBox.Show(NCCheck.errormessage);
                btnEmitirNC.IsEnabled = true;
                chkNCTribut.IsChecked = true;
                chkNCTribut.IsEnabled = true;
            }
            else
            {
                if (NCCheck.message != "")
                {
                System.Windows.Forms.MessageBox.Show(NCCheck.message);
                }
                label41.Visibility = Visibility.Visible;
                txtTotEfect.Visibility = Visibility.Visible;
                label10.Visibility = Visibility.Visible;
                DGDocDet.Visibility = Visibility.Visible;
                DGDocDet.ItemsSource = null;
                DGDocDet.Items.Clear();
                DGDocDet.ItemsSource = NCCheck.viapago;
                btnEmitirNC.IsEnabled = true;
            }
             label41.Visibility = Visibility.Visible;
             txtTotEfect.Visibility = Visibility.Visible;
             GC.Collect();
        }

        private void btnEmitirNC_Click(object sender, RoutedEventArgs e)
        {

           // List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> DocsAPagar = new List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();
            List<T_DOCUMENTOS_AUX> DocsAPagar = new List<T_DOCUMENTOS_AUX>();
                DocsAPagar.Clear();
                if (this.DGDocCabec.Items.Count > 0)
                {
                    for (int i = 0; i < DGDocCabec.Items.Count-1; i++)
                    {
                        if (i == 0)
                            DGDocCabec.Items.MoveCurrentToFirst();
                        {
                            DocsAPagar.Add(DGDocCabec.Items.CurrentItem as T_DOCUMENTOS_AUX);
                        }
                        DGDocCabec.Items.MoveCurrentToNext();
                    }
                }
                 List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS> partidaopen = new List<CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS>();

                for (int k = 0; k < DocsAPagar.Count; k++)
                {
                    if (DocsAPagar[k].ISSELECTED == true)
                    {
                        CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS partOpen = new CajaIndu.AppPersistencia.Class.NotasDeCredito.Estructura.T_DOCUMENTOS();
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
                //DGDocDet.ItemsSource = NCEmitir.viapago;
                btnEmitirNC.IsEnabled = false;
                ImpresionesDeDocumentosAutomaticas(NCEmitir.NumComprob, "X");
            }

         
            DGDocCabec.Visibility = Visibility.Collapsed;
            DGDocDet.Visibility = Visibility.Collapsed;
            label10.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility= Visibility.Collapsed;
            DGDocCabec.ItemsSource = null;
            DGDocCabec.Items.Clear();
            DGDocDet.ItemsSource = null;
            DGDocDet.Items.Clear();
            DGDocCabec.Visibility = Visibility.Collapsed;
            DGDocDet.Visibility = Visibility.Collapsed;
            label10.Visibility = Visibility.Collapsed;
            // }
             GC.Collect();
        }

        private void textBox22_TextChanged(object sender, TextChangedEventArgs e)
        {

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
               Total1= Convert.ToDouble(txtT1.Text);
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

           if (cmbMoneda.Text == "CLP")
           {
               txtTotalEfectivo.Text = FM.FormatoMonedaChilena(Convert.ToString(Total1 + Total5 + Total10 + Total50 + Total100 + Total500 + Total1000 + Total2000 + Total5000 + Total10000 + Total20000), "1");
           }
           else
           {
               txtTotalEfectivo.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total1+Total5+Total10+Total50+Total100+Total500+Total1000+Total2000+Total5000+Total10000+Total20000));
           }
        //txtTotalEfectivo.Text = Convert.ToString(Total1+Total5+Total10+Total50+Total100+Total500+Total1000+Total2000+Total5000+Total10000+Total20000);
        GC.Collect();   
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT1.Text = FM.FormatoMonedaChilena(Convert.ToString(Total1), "1");
                    }
                    else
                    {
                        txtT1.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total1));
                    }
                    //txtT1.Text = Convert.ToString(Total1);
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT5.Text = FM.FormatoMonedaChilena(Convert.ToString(Total5), "1");
                    }
                    else
                    {
                        txtT5.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total5));
                    }
                    //txtT5.Text = Convert.ToString(Total5);
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT10.Text = FM.FormatoMonedaChilena(Convert.ToString(Total10), "1");
                    }
                    else
                    {
                        txtT10.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total10));
                    }
                    //txtT10.Text = Convert.ToString(Total10);
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT50.Text = FM.FormatoMonedaChilena(Convert.ToString(Total50), "1");
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT100.Text = FM.FormatoMonedaChilena(Convert.ToString(Total100), "1");
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT500.Text = FM.FormatoMonedaChilena(Convert.ToString(Total500), "1");
                    }
                    else
                    {
                        txtT500.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total500));
                    }
                    //txtT500.Text = Convert.ToString(Total500);
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT1000.Text = FM.FormatoMonedaChilena(Convert.ToString(Total1000), "1");
                    }
                    else
                    {
                        txtT1000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total1000));
                    }
                    //txtT1000.Text = Convert.ToString(Total1000);
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT2000.Text = FM.FormatoMonedaChilena(Convert.ToString(Total2000), "1");
                    }
                    else
                    {
                        txtT2000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total2000));
                    }
                    //txtT2000.Text = Convert.ToString(Total2000);
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT5000.Text = FM.FormatoMonedaChilena(Convert.ToString(Total5000), "1");
                    }
                    else
                    {
                        txtT5000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total5000));
                    }
                    //txtT5000.Text = Convert.ToString(Total5000);
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
                    if (cmbMoneda.Text == "CLP")
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
                    if (cmbMoneda.Text == "CLP")
                    {
                        txtT20000.Text = FM.FormatoMonedaChilena(Convert.ToString(Total20000), "1");
                    }
                    else
                    {
                        txtT20000.Text = FM.FormatoMonedaExtranjera(Convert.ToString(Total20000));
                    }
                   //txtT20000.Text = Convert.ToString(Total20000);
                }
                else
                    txtC20000.Text= "0";
            }
            else
                System.Windows.MessageBox.Show("Ingrese un valor númerico entero");

            SumaEfectivoPorDenominacion();
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

        private void LimpiarEntradasDeDatos()
        {
            txtRuts.Text = "";
            txtDocum.Text = "";
            txtRUTReimp.Text = "";
            txtComprReimp.Text = "";
            txtRUTAn.Text = "";
            txtComprAn.Text = "";
            txtRUTAnV.Text = "";
            txtComprAnV.Text = "";
            txtRUTNC.Text = "";
            txtComprNC.Text = "";
            txtRut.Text = "";
            txtDocu.Text = "";
            txtArchivo.Text = "";
            txtRUTAnt.Text = "";
            txtDocum.Text = "";
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
                //List<T_DOCUMENTOS> partidaopen = new List<T_DOCUMENTOS>;
                
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

                    bool dalecandelanegro = false;
                    int validador = 0;
                   
                    if (DocsAPagar.Count > 0)
                    {
                        dalecandelanegro = true;
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
                   
                    if (dalecandelanegro == true)
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, ex.StackTrace);
                GC.Collect();
            }      
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
                    else
                    {
                        // ls_paramtCab.ACC = "1030";
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
                    switch (bapi_return2[i].TYPE)
                    {
                        case "E":
                            //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_stop.gif' />&nbsp;";
                            break;
                        case "I":
                            //cadImagen = "&nbsp;<img src='../../../Images/info.gif' />&nbsp;";
                            break;
                        case "W":
                            //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_warning.gif' />&nbsp;";
                            break;
                        case "S":
                            //cadImagen = "&nbsp;<img src='../../../Images/ico12_msg_success.gif' />&nbsp;";
                            break;
                    }
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

                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock7.Content), Convert.ToString(textBlock6.Content), Convert.ToString(textBlock8.Content), Mensaje);
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
                    //pagoanticipos.T_Retorno.Clear();

                }
                else
                {
                    System.Windows.MessageBox.Show("No se generó comprobante de pago");
                }

                //pagoanticipos.T_Retorno.Clear();
                //RECAUDA.objPag.Clear();
            }

            GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, ex.StackTrace);
                GC.Collect();
            }
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
            
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
            //RFC BUSQUEDA PARA GESTION DE DEPOSITOS
            GestionDeDepositos gdepot = new GestionDeDepositos();

            if (txtArqueo.Text != "")
            {

                gdepot.gestiondedepositos(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                    , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                    , Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content), LogApert[0].ID_REGISTRO,txtNumCierre.Text, txtArqueo.Text);

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
                    //DGViasPagoGD.ItemsSource = gdepot.vpgestiondepositos;
                    List<string> BancoDst = new List<string>();
                    List<string> CtaContable = new List<string>();
                    for (int i = 0; i < gdepot.BancoDeposito.Count; i++)
                    {
                        BancoDst.Add(gdepot.BancoDeposito[i].BANKN + "-" + gdepot.BancoDeposito[i].BANKL + "-" + gdepot.BancoDeposito[i].BANKA);
                        CtaContable.Add(gdepot.BancoDeposito[i].HKONT);
                    }
                    cmbBancoDest.ItemsSource = BancoDst;
                    cmbCuentaContable.ItemsSource = CtaContable;
                }
            }
        }
        //BOTON QUE AUTORIZA LAS ANULACIONES DE DOCUMENTOS POR PARTE DEL SUPER USUARIO 
        private void btnAutAnul_Click(object sender, RoutedEventArgs e)
        {           
            //RFC QUE MARCA LOS DOCUMENTOS A SER ANULADOS POR PARTE DEL SUPER USUARIO
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
                , Convert.ToString(lblPais.Content), txtRUTAnV.Text, txtComprAnV.Text, Convert.ToString(lblSociedad.Content),"A", Comprobante);
        
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

        private void cmbBancoDest_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btnDepositos.IsEnabled = true;
        }

        private void btnDepositos_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                     List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
                    for (int i = 1; i < DGLogApertura.Items.Count; i++)
                    {
                        if (i == 1)
                        {
                            DGLogApertura.Items.MoveCurrentToFirst();
                        }
                        if (DGLogApertura.Items.CurrentItem != null)
                        {
                            LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                        }
                        DGLogApertura.Items.MoveCurrentToNext();
                    }

                    DepositoProceso depotProcess = new DepositoProceso();

                    List<VIAS_PAGOGDAUX> Comprobante = new List<VIAS_PAGOGDAUX>();
                    Comprobante.Clear();
                        if (this.DGViasPagoGD.Items.Count > 0)
                        {
                            for (int i = 0; i < DGViasPagoGD.Items.Count-1; i++)
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
                                    //partOpen.ISSELECTED = false;
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
 
                        }
                        if (Validador == true)
                        {
                            depotProcess.depositoproceso(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                                , Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content), LogApert[0].ID_REGISTRO,txtNumCierre.Text, txtArqueo.Text,DPFechaDeposito.Text
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
                                    , Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content), LogApert[0].ID_REGISTRO,txtNumCierre.Text, txtArqueo.Text);

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
                                    ///////
                                    ///////
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
                                }
                            }

                               // DGViasPagoGD.ItemsSource = gdepot.vpgestiondepositos;
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
        }

        private void btnAnulCierre_Click(object sender, RoutedEventArgs e)
        {
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
            //***RFC cierre de Caja
            CierreCajaDefinitivo cierrecaja = new CierreCajaDefinitivo();
            cierrecaja.cierrecajadefinitivo(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), LogApert[0].ID_REGISTRO, txtTotalCaja.Text, txtDiferencia.Text, txtCommDif.Text, txtCommCierre.Text, LogApert[0].MONTO
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
                LimpiarElementosDeCierreDeCaja();
            }
            //Si existe ID de Cierre, bloquear botones de Menu salvo gestion de depositos

            GC.Collect();
           
        }
        //SELECCION DE CUENTA CONTABLE AUTOMATICA DE ACUERDO AL BANCO DESTINO EN GESTION DE DEPOSITOS
        private void cmbBancoDest_DropDownClosed(object sender, EventArgs e)
        {
            int posicion;

            posicion = cmbBancoDest.SelectedIndex;
            cmbCuentaContable.SelectedIndex = posicion;
            GC.Collect();

        }

        private void btnPreCierre_Click(object sender, RoutedEventArgs e)
        {
            if ((DPickDesde.Text != "") & (DPickHasta.Text != ""))
            {
            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }
            //RFC REPORTE DE CAJAS
            ReportesCaja reportcajas = new ReportesCaja();
            reportcajas.reportescaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                , Convert.ToString(lblPais.Content), Convert.ToString(lblSociedad.Content), LogApert[0].ID_REGISTRO, txtNumCierre.Text, "1");
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

        private void btnResumenCajas_Click(object sender, RoutedEventArgs e)
        {
             if ((DPickDesde.Text != "") & (DPickHasta.Text != ""))
            {

                List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
                for (int i = 1; i < DGLogApertura.Items.Count; i++)
                {
                    if (i == 1)
                    {
                        DGLogApertura.Items.MoveCurrentToFirst();
                    }
                    if (DGLogApertura.Items.CurrentItem != null)
                    {
                        LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                    }
                    DGLogApertura.Items.MoveCurrentToNext();
                }
                //RFC REPORTE DE CAJAS
                ReportesCaja reportcajas = new ReportesCaja();
                reportcajas.reportescaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                    , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                    , Convert.ToString(lblPais.Content), Convert.ToString(lblSociedad.Content), LogApert[0].ID_REGISTRO, txtNumCierre.Text, "3");
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

 
        private void btnResumenMovimientos_Click(object sender, RoutedEventArgs e)
        {
            if ((DPickDesde.Text != "") & (DPickHasta.Text != ""))
            {

                List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
                for (int i = 1; i < DGLogApertura.Items.Count; i++)
                {
                    if (i == 1)
                    {
                        DGLogApertura.Items.MoveCurrentToFirst();
                    }
                    if (DGLogApertura.Items.CurrentItem != null)
                    {
                        LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                    }
                    DGLogApertura.Items.MoveCurrentToNext();
                }
                //RFC REPORTE DE CAJAS
                ReportesCaja reportcajas = new ReportesCaja();
                reportcajas.reportescaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text
                    , txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content), DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content)
                    , Convert.ToString(lblPais.Content), Convert.ToString(lblSociedad.Content), LogApert[0].ID_REGISTRO, txtNumCierre.Text, "2");

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

        private void ImpresionReporteCaja(List<RENDICION_CAJA> ListRendicionCaja, List<RESUMEN_MENSUAL> ListResumenMensual, List<RESUMEN_CAJA> ListResumenCaja, string SociedadR, string Empresa, string Sucursal
            , string RUT ,string FechaDesde, string FechaHasta, string Tipo)
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
                    direct = direct + "InduLog\\" + "RendicionCaja" + fecha + ".pdf";
                    watermarkedFile = appRootDir + "InduLog\\" + "RendicionCaja" + fecha + "-Nuevo.Text.pdf";

                }
                if (Tipo == "2")
                {
                    direct = direct + "InduLog\\" + "ResumenMensualMovimientos" + fecha + ".pdf";                 
                    watermarkedFile = appRootDir + "InduLog\\" + "ResumenMensualMovimientos" + fecha + "-Nuevo.Text.pdf";
                }
                if (Tipo == "3")
                {
                    direct = direct + "InduLog\\" + "ResumenCaja" + fecha + ".pdf";
                    watermarkedFile =  appRootDir + "InduLog\\" + "ResumenCaja" + fecha + "-Nuevo.Text.pdf";
                }
            //MARCA DE AGUA PARA DOCUMENTOS DE PRUEBA O SIN VALIDEZ
			string watermarkText = "Documento No válido";
            string Cajero = "";
           
            using (FileStream fs = new FileStream(direct, FileMode.Create, FileAccess.Write, FileShare.None))
               
            using (Document pdfcommande = new Document(PageSize.LETTER.Rotate(), 20f,20f,100f,100f))
			
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
                    //INFORME DE RENDICION DE CAJA
                    //LLENADO DE LA GRILLA
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(MONTO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow4.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow4);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
                                        {
                                            MonedaFormateada = FM.FormatoMonedaChilena(ViasPago[k].MONTO_EFEC, "2");
                                            PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            table2.AddCell(cellrow6);
                                        }
                                        else
                                        {
                                            MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                            PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                            table2.AddCell(cellrow6);
                                        }
                                        MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                    }
                                    else
                                    {
                                        //PdfPCell cellrow6 = new PdfPCell(new Phrase(Convert.ToString(MONTO_EFEC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow6);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow8 = new PdfPCell(new Phrase(Convert.ToString(MONTO_DIA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow8);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(MONTO_FECHA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow9.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow9);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow10 = new PdfPCell(new Phrase(Convert.ToString(MONTO_TRANSF), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow10);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(MONTO_VALE_V), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow11);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow12 = new PdfPCell(new Phrase(Convert.ToString(MONTO_DEP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow12);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow13 = new PdfPCell(new Phrase(Convert.ToString(MONTO_TARJ), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow13);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow14 = new PdfPCell(new Phrase(Convert.ToString(MONTO_FINANC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow14);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow15 = new PdfPCell(new Phrase(Convert.ToString(MONTO_APP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow15);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow16 = new PdfPCell(new Phrase(Convert.ToString(MONTO_CREDITO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow16);

                                        if (cmbMoneda.Text == "CLP")
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
                        else {
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow10 = new PdfPCell(new Phrase(Convert.ToString(MONTO_EFEC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow10);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(MONTO_DIA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow11);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow12 = new PdfPCell(new Phrase(Convert.ToString(MONTO_FECHA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow12);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow13 = new PdfPCell(new Phrase(Convert.ToString(MONTO_TRANSF), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow13);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow14 = new PdfPCell(new Phrase(Convert.ToString(MONTO_VALE_V), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow14);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow15 = new PdfPCell(new Phrase(Convert.ToString(MONTO_DEP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow15);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow16 = new PdfPCell(new Phrase(Convert.ToString(MONTO_TARJ), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow16);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow17 = new PdfPCell(new Phrase(Convert.ToString(MONTO_FINANC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow17);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow18 = new PdfPCell(new Phrase(Convert.ToString(MONTO_APP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow18);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow19 = new PdfPCell(new Phrase(Convert.ToString(MONTO_CREDITO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow19);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow20 = new PdfPCell(new Phrase(Convert.ToString(TOTAL_CAJERO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow20);
                                        if (cmbMoneda.Text == "CLP")
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
                        else {
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow10 = new PdfPCell(new Phrase(Convert.ToString(MONTO_EFEC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow10);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(MONTO_DIA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow11);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow12 = new PdfPCell(new Phrase(Convert.ToString(MONTO_FECHA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow12);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow13 = new PdfPCell(new Phrase(Convert.ToString(MONTO_TRANSF), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow13);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow14 = new PdfPCell(new Phrase(Convert.ToString(MONTO_VALE_V), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow14);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow15 = new PdfPCell(new Phrase(Convert.ToString(MONTO_DEP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow15);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow16 = new PdfPCell(new Phrase(Convert.ToString(MONTO_TARJ), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow16);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow17 = new PdfPCell(new Phrase(Convert.ToString(MONTO_FINANC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow17);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow18 = new PdfPCell(new Phrase(Convert.ToString(MONTO_APP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow18);
                                        if (cmbMoneda.Text == "CLP")
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
                                        if (cmbMoneda.Text == "CLP")
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
                                        //PdfPCell cellrow19 = new PdfPCell(new Phrase(Convert.ToString(MONTO_CREDITO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                        //cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                        //table2.AddCell(cellrow19);
                                        if (cmbMoneda.Text == "CLP")
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
                        else {
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
                // Creating watermark on a separate layer
                // Creating iTextSharp.text.pdf.PdfReader object to read the Existing PDF Document produced by 1 no.
                string direct2 = Convert.ToString(System.IO.Path.GetTempPath());

                PdfReader reader1 = new PdfReader(direct);
                using (FileStream fs = new FileStream(watermarkedFile, FileMode.Create, FileAccess.Write, FileShare.None))
                // Creating iTextSharp.text.pdf.PdfStamper object to write Data from iTextSharp.text.pdf.PdfReader object to FileStream object
                using (PdfStamper stamper = new PdfStamper(reader1, fs))
                {
                    // Getting total number of pages of the Existing Document
                    int pageCount = reader1.NumberOfPages;

                    // Create New Layer for Watermark
                    PdfLayer layer = new PdfLayer("WatermarkLayer", stamper.Writer);

                    //PdfLayer layer2 = new PdfLayer("Paginacion",tablex);
                    // Loop through each Page
                    for (int i = 1; i <= pageCount; i++)
                    {
                        string PaginaActual = Convert.ToString(i);
                        string PaginasTotales = Convert.ToString(pageCount);
                        // Getting the Page Size
                        iTextSharp.text.Rectangle rect = reader1.GetPageSize(i);
                        rect.Rotate();
                        if (txtMandante.Text != "300")
                        {
                            // Get the ContentByte object
                            PdfContentByte cb = stamper.GetUnderContent(i);
                            // Tell the cb that the next commands should be "bound" to this new layer
                            cb.BeginLayer(layer);
                            cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 50);
                            PdfGState gState = new PdfGState();
                            gState.FillOpacity = 0.25f;
                            cb.SetGState(gState);
                            cb.SetColorFill(BaseColor.BLACK);
                            cb.BeginText();
                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, watermarkText, rect.Width / 2, rect.Height / 2, 45f);
                            cb.EndText();
                            // Close the layer
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
                        //cb2.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Paginas, 10, 10, 0f);
                        cb2.EndText();
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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
                        // Close the layer
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


        private void ReportesCaja_Click(object sender, RoutedEventArgs e)
        {
            GBInicio.Visibility = Visibility.Collapsed;
            GBMonitor.Visibility = Visibility.Collapsed;
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBPagoDocs.Visibility = Visibility.Collapsed;
            GBInicio.Visibility = Visibility.Collapsed;
            GBAnulacion.Visibility = Visibility.Collapsed;
            GBReimpresion.Visibility = Visibility.Collapsed;
            GBViasPago.Visibility = Visibility.Collapsed;
            GBDocsAPagar.Visibility = Visibility.Collapsed;
            GBDetalleDocs.Visibility = Visibility.Collapsed;
            GBEmisionNC.Visibility = Visibility.Collapsed;
            GBRendicion.Visibility = Visibility.Collapsed;
            GBrecauda.Visibility = Visibility.Collapsed;
            GBDocs.Visibility = Visibility.Collapsed;
            GBResumenCaja.Visibility = Visibility.Collapsed;
            GBDetEfectivo.Visibility = Visibility.Collapsed;
            GBCierreCaja.Visibility = Visibility.Collapsed;
            GBCommentCierre.Visibility = Visibility.Collapsed;
            GBAutorizacionVehiculos.Visibility = Visibility.Collapsed;
            GBGestionDeBancos.Visibility = Visibility.Collapsed;
            GBReportes.Visibility = Visibility.Visible;
            LimpiarViasDePago();
            LimpiarElementosDeCierreDeCaja();
            LimpiarEntradasDeDatos();

            List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
            for (int i = 1; i < DGLogApertura.Items.Count; i++)
            {
                if (i == 1)
                {
                    DGLogApertura.Items.MoveCurrentToFirst();
                }
                if (DGLogApertura.Items.CurrentItem != null)
                {
                    LogApert.Add(DGLogApertura.Items.CurrentItem as LOG_APERTURA);
                }
                DGLogApertura.Items.MoveCurrentToNext();
            }


            //RFC Rendicion Caja
            RendicionCaja rendicioncaja = new RendicionCaja();
            rendicioncaja.rendicioncaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text
                , txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, Convert.ToString(textBlock6.Content)
                , DPickDesde.Text, DPickHasta.Text, Convert.ToString(textBlock7.Content), Convert.ToString(lblPais.Content)
                , Convert.ToString(lblSociedad.Content), LogApert[0].ID_REGISTRO, "0000000000", "0000000000", LogApert[0].MONEDA);


            if (rendicioncaja.detalle_rend.Count > 0)
            {
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("No existen datos para el informe de rendición");
            }
            if ((rendicioncaja.id_arqueo != "0000000000") & (rendicioncaja.id_cierre != "0000000000"))
            {
                //btnGestionDep.IsEnabled = false;
                txtArqueo.Text = rendicioncaja.id_arqueo;
                txtNumCierre.Text = rendicioncaja.id_cierre;
                //txtDiferencia.Text = "0";
                System.Windows.Forms.MessageBox.Show("Esta caja posee un proceso de arqueo y cierre previo");
            }
            else
            {
                if (rendicioncaja.id_arqueo != "0000000000")
                {
                    //btnGestionDep.IsEnabled = false;
                    txtArqueo.Text = rendicioncaja.id_arqueo;
                    txtNumCierre.Text = rendicioncaja.id_cierre;
                    //txtDiferencia.Text = "0";
                    System.Windows.Forms.MessageBox.Show("Esta caja posee un proceso de arqueo previo");
                }
            }
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

        private void ExportaDataToExcel(List<RENDICION_CAJA> ListRendicionCaja, List<RESUMEN_MENSUAL> ListResumenMensual, List<RESUMEN_CAJA> ListResumenCaja, string SociedadR, string Empresa, string Sucursal
            , string RUT, string FechaDesde, string FechaHasta, string Tipo, string Caja)
        {

            try
            {
                //Microsoft.Office.Interop.Excel.Application Excel;
                Microsoft.Office.Interop.Excel.Application xlApp;// = new Microsoft.Office.Interop.Excel.ApplicationClass();
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                try
                {
                    // Agregamos Los datos que queremos agregar
                    //xlWorkSheet.Cells["A", "3"].Value = Empresa;
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
                        // xlWorkSheet.Columns["C"].Width = 0;
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
                            //xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = ViasPago[k].MONTO;
                            if (k != ViasPago.Count)
                            {
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_EFEC;

                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_DIA;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_FECHA;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_TRANSF;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_VALE_V;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = ViasPago[k].;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_TARJ;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_FINANC;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_APP;
                                if (cmbMoneda.Text == "CLP")
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
                                //xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = ViasPago[k].MONTO_CREDITO;
                                if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        // xlWorkSheet.Columns["C"].Width = 0;
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                                if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        //if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        // xlWorkSheet.Columns["C"].Width = 0;
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                            if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                        if (cmbMoneda.Text == "CLP")
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
                    //xlWorkBook.SaveAs("Rendicion" + "-" + Caja + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                if (Tipo == "2")
                {
                    xlApp.Visible = true;
                    //xlWorkBook.SaveAs("ResumenMovimientos" + "-" + Caja + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }
                if (Tipo == "3")
                {
                    xlApp.Visible = true;
                    //xlWorkBook.SaveAs("ResumenCaja" + "-" + Caja + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                }

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                System.Windows.Forms.MessageBox.Show("Archivo Excel creado");

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }

        }

        private void frame2_Navigated(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {

        }

        private void frame1_Navigated(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {

        }

        private void btnInformePreCierre_Click(object sender, RoutedEventArgs e)
        {
            ImpresionInformePreCierre();
        }

        private void chkNCTribut_Checked(object sender, RoutedEventArgs e)
        {
            // btnEmitirNC.IsEnabled = true;
        }

        private void chkNCTribut_Unchecked(object sender, RoutedEventArgs e)
        {
            // btnEmitirNC.IsEnabled = false;
        }

    }
}