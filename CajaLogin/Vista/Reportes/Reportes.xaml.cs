using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura;
using System.Windows.Threading;
using CajaIndigo.AppPersistencia.Class.ReportesCaja;
using CajaIndigo.AppPersistencia.Class.ReportesCaja.Estructura;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using CajaIndigo.AppPersistencia.Class.BloquearCaja;

namespace CajaIndigo.Vista.Reportes
{
    public partial class Reportes : Window
    {
        List<LOG_APERTURA> logApertura = new List<LOG_APERTURA>();
        List<LOG_APERTURA> logApertura2 = new List<LOG_APERTURA>();
        DispatcherTimer timer = new DispatcherTimer();

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
        string Valor2 = string.Empty;
        string moneda = string.Empty;

        Vista.PagoDocumento.PagoDocumento PagDocum;
        Vista.NotaCredito.NotaCredito NotaCredit;
        Vista.Anulacion.Anulacion Anula;
        Vista.Reimpresion.Reimpresion Reimp;
        Vista.Vehiculos.Vehiculo Vehi;
        Vista.CierreCaja.CierreCaja CierCaja;
        Vista.Reportes.Reportes Reporte;

        public Reportes(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, string moneda, string pais, double monto, string IdSistema, string Instancia, string mandante, string SapRouter, string server, string idioma, List<LOG_APERTURA> logApertura)
            {
                try
                {
                    InitializeComponent();
                    int test = 0;
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
                    //PaisCaja.Text = pais;
                    DateTime result = DateTime.Today;
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
                Reporte = new Reportes(UserCaja, PassCaja, UserCaja, IdCaja, NombCaja, SociedadCaja, MonedCaja, PaisCja, Convert.ToDouble(Monto), IdSistema, Instancia, mandante, SapRouter, server, idioma, logApertura2);
                Reporte.Show();
                this.Hide();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void bt_recaudacion(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            CargarDatos();
        }

        private void ReportesCaja_Click(object sender, RoutedEventArgs e)
        {
            CargarDatos();
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

        private void ImpresionReporteCaja(List<RENDICION_CAJA> ListRendicionCaja, List<RESUMEN_MENSUAL> ListResumenMensual, List<RESUMEN_CAJA> ListResumenCaja, string SociedadR, string Empresa, string Sucursal
                  , string RUT, string FechaDesde, string FechaHasta, string Tipo)
        {
            try
            {
                //string fecha = Convert.ToString(DateTime.Now);
                //fecha = fecha.Replace(" ", "-");
                //fecha = fecha.Replace(":", "-");
                //string appRootDir = Convert.ToString(System.IO.Path.GetTempPath());
                //string watermarkedFile = "";
                //string direct = Convert.ToString(System.IO.Path.GetTempPath());
                ////string direct = string.Empty;

                string watermarkText = "Documento No válido";
                string Cajero = "";

                string fecha = Convert.ToString(DateTime.Now.ToString("hh:mm:ss.F"));
                fecha = fecha.Replace(" ", "-");
                fecha = fecha.Replace(":", "-");
                string Directorio = System.IO.Path.GetTempPath();
                Directorio = Directorio + "InchcapeLog\\";
                string watermarkedFile = "";

                Document documento = new Document(PageSize.LEGAL.Rotate(), 10f, 10f, 100f, 100f);

                if (Tipo == "1")
                {
                    Directorio = Directorio + "RendicionCaja" + fecha + ".pdf";
                    watermarkedFile = Directorio + "-Nuevo.Text.pdf";
                }
                if (Tipo == "2")
                {
                    Directorio = Directorio + "ResumenMensualMovimientos" + fecha + ".pdf";
                    watermarkedFile = Directorio + "-Nuevo.Text.pdf";
                }
                if (Tipo == "3")
                {
                    Directorio = Directorio + "ResumenCaja" + fecha + ".pdf";
                    watermarkedFile = Directorio + "-Nuevo.Text.pdf";
                }


                //using (FileStream fs = new FileStream(direct, FileMode.Create, FileAccess.Write, FileShare.None))

                //using (Document pdfcommande = new Document(PageSize.LETTER.Rotate(), 20f, 20f, 40f, 40f))

                //using (PdfWriter writer = PdfWriter.GetInstance(pdfcommande, fs))
                PdfWriter.GetInstance(documento, new FileStream(Directorio, FileMode.Create, FileAccess.Write, FileShare.None));
                {
                    //{
                        string direc = Directorio;
                        try
                    {
                        //pdfcommande.Open();

                        //pdfcommande.NewPage();
                        documento.Open();

                        documento.NewPage();
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
                        itxtTitulo.IndentationLeft = 470;
                        itxtTitulo.Font.Size = 10;
                        itxtTitulo.Font.SetStyle(1);
                        itxtTitulo.Font.SetFamily("Courier");
                        itxtTitulo.SpacingBefore = 30f;
                        itxtTitulo.SpacingAfter = 5f;
                        documento.Add(itxtTitulo);
                        texto = "Desde: " + FechaDesde + "           " + "Hasta: " + FechaHasta;
                        iTextSharp.text.Paragraph itxtfechDesdeHasta = new iTextSharp.text.Paragraph(texto);
                        itxtfechDesdeHasta.IndentationLeft = 450;
                        itxtfechDesdeHasta.Font.Size = 9;
                        itxtfechDesdeHasta.Font.SetFamily("Courier");
                        itxtfechDesdeHasta.SpacingAfter = 5f;
                        documento.Add(itxtfechDesdeHasta);
                        //Datos Caja
                        texto = "Id Caja: " + Convert.ToString(textBlock6.Content);
                        iTextSharp.text.Paragraph itxtIngreso = new iTextSharp.text.Paragraph(texto);
                        itxtIngreso.IndentationLeft = 1;
                        itxtIngreso.Font.Size = 9;
                        itxtIngreso.Font.SetFamily("Courier");
                        documento.Add(itxtIngreso);
                        //Datos de nota venta
                        texto = "Sucursal: " + Convert.ToString(textBlock8.Content);
                        iTextSharp.text.Paragraph itxtNotaVta = new iTextSharp.text.Paragraph(texto);
                        itxtNotaVta.IndentationLeft = 1;
                        itxtNotaVta.Font.Size = 9;
                        itxtNotaVta.Font.SetFamily("Courier");
                        itxtNotaVta.SpacingAfter = 10;
                        documento.Add(itxtNotaVta);

                        PdfPTable table2;
                        PdfPTable table3;

                        double MONTO = 0;
                        double MONTO2 = 0;
                        double MONTO3 = 0;
                        double MONTO4 = 0;
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
                                table2.TotalWidth = 990f;
                                string tableheight = Convert.ToString(table2.TotalHeight);
                                table2.LockedWidth = true;
                                table2.HeaderRows = 2;
                                table2.SpacingAfter = 5F;


                                //ANCHO DE LAS COLUMNAS
                                float[] widths = new float[] { 20f, 45f, 0f, 50f, 50f, 50f, 50f, 50f, 40f, 40f, 40f, 40f, 50f, 50f, 50f, 50f, 50f, 50f, 50f, 50f };
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
                                    try
                                    {
                                        //HASTA LA POSICION MAXIMA DE LA LISTA - 1, SE MUESTRAN LOS DATOS DE LA LISTA.
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].N_VENTA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        //EN LA POSICION MAXIMA, O ES VACIO, O SE COLOCA ALGUN TITULO O SE MUESTRAN LOS TOTALES DE LAS COLUMNAS CON MONTOS.
                                        else
                                        {
                                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow1);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DOC_TRIB), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow2);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].CAJERO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow3);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FEC_EMI), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow4);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow4 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow4);
                                        }
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].FEC_VENC), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow5);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow5 = new PdfPCell(new Phrase(Convert.ToString("Totales"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow5);
                                        }
                                        //MONTO
                                        //HASTA LA POSICION MAXIMA DE LA LISTA - 1, SE MUESTRAN LOS DATOS DE LA LISTA.
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO);
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO);
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }
                                           
                                                MONTO = MONTO + Convert.ToDouble(ViasPago[k].MONTO);
      
                                        }
                                        //EN LA POSICION MAXIMA, O ES VACIO, O SE COLOCA ALGUN TITULO O SE MUESTRAN LOS TOTALES DE LAS COLUMNAS CON MONTOS.
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO));
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO));
                                                PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow6.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow6);
                                            }
                                            //MONTO = MONTO + Convert.ToDouble(ViasPago[k].MONTO);
                                        }
                                        //NAME1
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NAME1), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow7);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow7 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow7);
                                        }
                                        //MONTO_EFEC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_EFEC);
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                                PdfPCell cellrow8 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow8.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow8);
                                            }
                                            if (ViasPago[k].MONEDA == "CLP")
                                            {
                                                MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                            }
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_EFEC));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_EFEC));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }                                 
                                        }
                                        // MONEDA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONEDA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow9);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow9);
                                        }
                                        //NUM_CHEQUE
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NUM_CHEQUE), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow11);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow11 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow11);
                                        }
                                        //MONTO_DIA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                        }
                                        //MONTO_FECHA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                        }
                                        //MONTO_TRANSF
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                        }
                                        //MONTO_VALE_V
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                        }
                                        //MONTO_DEP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                        }
                                        //MONTO_TARJ
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                        }
                                        //MONTO_FINANC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                        }
                                        //MONTO_APP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                        }
                                        //MONTO_CREDITO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);
                                            }
                                            MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);
                                            }
                                        }
                                        //DOC_SAP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow22 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DOC_SAP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow22);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow22 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 7f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow22);
                                        }
                                        // LLENAS EFECTIVO OTRAS MONEDAS
                                        if (ViasPago[k].MONEDA == "CLP")
                                        {
                                            MonedaFormateada = FM.FormatoMoneda(Convert.ToString(ViasPago[k].MONTO_EFEC));
                                            MONTO4 = MONTO4 + Convert.ToDouble(MonedaFormateada);
                                        }
                                        if (ViasPago[k].MONEDA == "USD")
                                        {
                                            MONTO2 = MONTO2 + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                        if (ViasPago[k].MONEDA == "EUR")
                                        {
                                            MONTO3 = MONTO3 + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.Write(ex.Message, ex.StackTrace);
                                    }
                                }
                                //Tabla Totales                               
                                PdfPTable tabla = new PdfPTable(2);
                                PdfPTable tabla3 = new PdfPTable(1);
                                tabla.TotalWidth = 775f;
                                tabla.HorizontalAlignment = Element.ALIGN_LEFT;
                                tabla.WidthPercentage = 30.0f;
                                tabla3.TotalWidth = 775f;
                                tabla3.HorizontalAlignment = Element.ALIGN_LEFT;
                                tabla3.WidthPercentage = 30.0f;
                                float[] widths2 = new float[] { 2f, 2f };
                                tabla.SetWidths(widths2);
                                for (int i = 0; i < 1; i++)
                                {
                                    PdfPCell CellTitulo0 = new PdfPCell(new Phrase(Convert.ToString("Totales en efectivo caja"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla3.AddCell(CellTitulo0);
                                    PdfPCell Celltitulo = new PdfPCell(new Phrase(Convert.ToString("Total CLP: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo);
                                    PdfPCell CellPeso = new PdfPCell(new Phrase((Convert.ToString(MONTO4)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellPeso);
                                    PdfPCell Celltitulo2 = new PdfPCell(new Phrase(Convert.ToString("Total USD: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo2);
                                    PdfPCell CellUSD = new PdfPCell(new Phrase((Convert.ToString(MONTO2)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellUSD);
                                    PdfPCell Celltitulo3 = new PdfPCell(new Phrase(Convert.ToString("Total EUR: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo3);
                                    PdfPCell CellEUR = new PdfPCell(new Phrase((Convert.ToString(MONTO3)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellEUR);
                                }
                                documento.Add(table2);
                                documento.Add(tabla3);
                                documento.Add(tabla);
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
                                float[] widths = new float[] { 30f, 100f, 30f, 100f, 40f, 30f, 80f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f, 70f };
                                table2.SetWidths(widths);
                                //int factor = 1;
                                PdfPCell cell2 = new PdfPCell(new Phrase(Convert.ToString(label11.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                                cell2.Padding = 10f;
                                cell2.Colspan = DGResumenMovimientosRep.Columns.Count;
                                cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                                table2.AddCell(cell2);
                                foreach (DataGridColumn column in DGResumenMovimientosRep.Columns)
                                {
                                    table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 4f, iTextSharp.text.Font.NORMAL)));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].TOTAL_MOV);
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
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(TOTAL_MOV));
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

                                        // MONEDA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].MONEDA), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow9);
                                        }
                                        else
                                        {
                                            PdfPCell cellrow9 = new PdfPCell(new Phrase(Convert.ToString(" "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 6f, iTextSharp.text.Font.NORMAL)));
                                            table2.AddCell(cellrow9);
                                        }

                                        //TOTAL_INGR
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].TOTAL_INGR);
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_INGR);
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            TOTAL_INGR = TOTAL_INGR + Convert.ToDouble(ViasPago[k].TOTAL_INGR);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(TOTAL_INGR));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_INGR));
                                                PdfPCell cellrow10 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow10.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow10);
                                            }
                                        }
                                        //MONTO_EFEC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_EFEC);
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_EFEC));
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                            else
                                            {
                                                PdfPCell cellrow11 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow11.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow11);
                                            }
                                        }
                                        //MONTO_DIA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                                                PdfPCell cellrow12 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow12.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow12);
                                            }
                                        }
                                        //MONTO_FECHA
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                                                PdfPCell cellrow13 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow13.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow13);
                                            }
                                        }
                                        //MONTO_TRANSF
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                                                PdfPCell cellrow14 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow14.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow14);
                                            }
                                        }
                                        //MONTO_VALE_V
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                                                PdfPCell cellrow15 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow15.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow15);
                                            }
                                        }
                                        //MONTO_DEP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                                                PdfPCell cellrow16 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow16.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow16);
                                            }
                                        }
                                        //MONTO_TARJ
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                                                PdfPCell cellrow17 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow17.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow17);
                                            }
                                        }
                                        //MONTO_FINANC
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                                                PdfPCell cellrow18 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow18.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow18);
                                            }
                                        }
                                        //MONTO_APP
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                                                PdfPCell cellrow19 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow19.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow19);
                                            }
                                        }
                                        //MONTO_CREDITO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                            MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                                                PdfPCell cellrow20 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow20.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow20);
                                            }
                                        }
                                        //TOTAL_CAJERO
                                        if (k != ViasPago.Count - 1)
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].TOTAL_CAJERO);
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);

                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_CAJERO);
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);
                                            }
                                            TOTAL_CAJERO = TOTAL_CAJERO + Convert.ToDouble(ViasPago[k].TOTAL_CAJERO);
                                        }
                                        else
                                        {
                                            if (logApertura2[0].MONEDA == "CLP")
                                            {
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(TOTAL_CAJERO));
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);

                                            }
                                            else
                                            {
                                                MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_CAJERO));
                                                PdfPCell cellrow21 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 5f, iTextSharp.text.Font.NORMAL)));
                                                cellrow21.HorizontalAlignment = Element.ALIGN_RIGHT;
                                                table2.AddCell(cellrow21);
                                            }
                                        }
                                        // LLENAS EFECTIVO OTRAS MONEDAS
                                        if (ViasPago[k].MONEDA == "CLP")
                                        {
                                            MonedaFormateada = FM.FormatoMoneda(Convert.ToString(ViasPago[k].MONTO_EFEC));
                                            MONTO4 = MONTO4 + Convert.ToDouble(MonedaFormateada);
                                        }
                                        if (ViasPago[k].MONEDA == "USD")
                                        {
                                            MONTO2 = MONTO2 + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                        if (ViasPago[k].MONEDA == "EUR")
                                        {
                                            MONTO3 = MONTO3 + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.Write(ex.Message, ex.StackTrace);
                                    }

                                }
                                //Tabla Totales                               
                                PdfPTable tabla = new PdfPTable(2);
                                PdfPTable tabla3 = new PdfPTable(1);
                                tabla.TotalWidth = 775f;
                                tabla.HorizontalAlignment = Element.ALIGN_LEFT;
                                tabla.WidthPercentage = 30.0f;
                                tabla3.TotalWidth = 775f;
                                tabla3.HorizontalAlignment = Element.ALIGN_LEFT;
                                tabla3.WidthPercentage = 30.0f;
                                float[] widths2 = new float[] { 2f, 2f };
                                tabla.SetWidths(widths2);
                                for (int i = 0; i < 1; i++)
                                {
                                    PdfPCell CellTitulo0 = new PdfPCell(new Phrase(Convert.ToString("Totales en efectivo caja"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla3.AddCell(CellTitulo0);
                                    PdfPCell Celltitulo = new PdfPCell(new Phrase(Convert.ToString("Total CLP: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo);
                                    PdfPCell CellPeso = new PdfPCell(new Phrase((Convert.ToString(MONTO4)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellPeso);
                                    PdfPCell Celltitulo2 = new PdfPCell(new Phrase(Convert.ToString("Total USD: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo2);
                                    PdfPCell CellUSD = new PdfPCell(new Phrase((Convert.ToString(MONTO2)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellUSD);
                                    PdfPCell Celltitulo3 = new PdfPCell(new Phrase(Convert.ToString("Total EUR: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo3);
                                    PdfPCell CellEUR = new PdfPCell(new Phrase((Convert.ToString(MONTO3)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellEUR);
                                }
                                documento.Add(table2);
                                documento.Add(tabla3);
                                documento.Add(tabla);
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
                                float[] widths = new float[] { 30f, 100f, 30f, 100f, 60f, 60f, 60f, 60f, 60f, 60f, 60f, 60f, 60f, 60f};
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
                                    table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 4f, iTextSharp.text.Font.NORMAL)));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_EFEC);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_EFEC));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_DIA);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_DIA));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_FECHA);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_FECHA));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_TRANSF);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_TRANSF));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_VALE_V);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_VALE_V));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_DEP);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_DEP));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_TARJ);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_TARJ));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_FINANC);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_FINANC));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_APP);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_APP));
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
                                                MonedaFormateada = FM.FormatoMoneda(ViasPago[k].MONTO_CREDITO);
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
                                                MonedaFormateada = FM.FormatoMoneda(Convert.ToString(MONTO_CREDITO));
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
                                //Tabla Totales                               
                                PdfPTable tabla = new PdfPTable(2);
                                PdfPTable tabla3 = new PdfPTable(1);
                                tabla.TotalWidth = 775f;
                                tabla.HorizontalAlignment = Element.ALIGN_LEFT;
                                tabla.WidthPercentage = 30.0f;
                                tabla3.TotalWidth = 775f;
                                tabla3.HorizontalAlignment = Element.ALIGN_LEFT;
                                tabla3.WidthPercentage = 30.0f;
                                float[] widths2 = new float[] { 2f, 2f };
                                tabla.SetWidths(widths2);
                                for (int i = 0; i < 1; i++)
                                {
                                    PdfPCell CellTitulo0 = new PdfPCell(new Phrase(Convert.ToString("Totales en efectivo caja"), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla3.AddCell(CellTitulo0);
                                    PdfPCell Celltitulo = new PdfPCell(new Phrase(Convert.ToString("Total CLP: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo);
                                    PdfPCell CellPeso = new PdfPCell(new Phrase((Convert.ToString(MONTO)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellPeso);
                                    PdfPCell Celltitulo2 = new PdfPCell(new Phrase(Convert.ToString("Total USD: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo2);
                                    PdfPCell CellUSD = new PdfPCell(new Phrase((Convert.ToString(MONTO2)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellUSD);
                                    PdfPCell Celltitulo3 = new PdfPCell(new Phrase(Convert.ToString("Total EUR: "), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(Celltitulo3);
                                    PdfPCell CellEUR = new PdfPCell(new Phrase((Convert.ToString(MONTO3)), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 12f, iTextSharp.text.Font.BOLD)));
                                    tabla.AddCell(CellEUR);
                                }
                                documento.Add(table2);
                                documento.Add(tabla3);
                                documento.Add(tabla);
                            }
                            else
                            {
                                System.Windows.Forms.MessageBox.Show("No Existen datos Para el intervalo seleccionado");
                            }
                        }
                        documento.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.Message, ex.StackTrace);
                    }
                }

                try
                {
                    string direct2 = Convert.ToString(System.IO.Path.GetTempPath());

                    PdfReader reader1 = new PdfReader(Directorio);
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
                            cb2.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Paginas, rect.Width - (rect.Width - 870), rect.Height -450 , 0f);                        
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
                            cb3.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Empresa, rect.Width - (rect.Width - 8), rect.Height - 450, 0f);
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
                            cb4a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Fecha:", rect.Width - (rect.Width - 870), rect.Height - 460, 0f);
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
                            cb4.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, String.Format("{0:dd/MM/yyyy}", FechaAFormatear), rect.Width - (rect.Width - 1000), rect.Height - 460, 0f);
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
                            cb5.ShowTextAligned(PdfContentByte.ALIGN_LEFT, RUT, rect.Width - (rect.Width - 8), rect.Height - 460, 0f);
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
                            cb6.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, String.Format("{0:HH:mm:ss}", FechaAFormatear), rect.Width - (rect.Width - 1000), rect.Height - 470, 0f);
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
                            cb6a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Hora:", rect.Width - (rect.Width - 870), rect.Height - 470, 0f);
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
                            cb7.ShowTextAligned(PdfContentByte.ALIGN_LEFT, SociedadR, rect.Width - (rect.Width - 10), rect.Height - 320, 0f);
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
                            cb8.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, Convert.ToString(textBlock7.Content), rect.Width - (rect.Width - 1000), rect.Height - 480, 0f);
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
                            cb8a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Usuario:", rect.Width - (rect.Width - 870), rect.Height - 480, 0f);
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
                                cb9.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, Cajero, rect.Width - (rect.Width - 1000), rect.Height - 490, 0f);
                                cb9.EndText();
                                cb9.EndLayer();

                                //CAJERO LABEL
                                PdfContentByte cb9a = stamper.GetUnderContent(i);
                                cb9a.BeginLayer(layer);
                                cb9a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb9a.SetGState(gState2);
                                cb9a.SetColorFill(BaseColor.BLACK);
                                cb9a.BeginText();
                                cb9a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Cajero:", rect.Width - (rect.Width - 870), rect.Height - 490, 0f);
                                cb9a.EndText();
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
                {
                    // Agregamos Los datos que queremos agreg
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
                xlWorkSheet.Range["B7"].Value = String.Format("{0:dd/MM/yyyy}", FechaDesde);
                xlWorkSheet.Range["A8"].Value = "Hasta:";
                xlWorkSheet.Range["B8"].Value = String.Format("{0:dd/MM/yyyy}", FechaHasta);

                int i = 0;
                int j = 0;

                DGRendicionCajaRep.ItemsSource = ListRendicionCaja;
                DGResumenCajasRep.ItemsSource = ListResumenCaja;
                DGResumenMovimientosRep.ItemsSource = ListResumenMensual;

                double MONTO = 0;
                double MONTO2 = 0;
                double MONTO3 = 0;
                double MONTO4 = 0;
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
                        xlWorkSheet.Range["I10"].Value = "Moneda";
                        xlWorkSheet.Range["J10"].Value = "N° doc.";
                        xlWorkSheet.Range["K10"].Value = "Chq. al día";
                        xlWorkSheet.Range["L10"].Value = "Chq. a fecha";
                        xlWorkSheet.Range["M10"].Value = "Transferencia";
                        xlWorkSheet.Range["N10"].Value = "V. vista";
                        xlWorkSheet.Range["O10"].Value = "Depósitos";
                        xlWorkSheet.Range["P10"].Value = "Tarjetas";
                        xlWorkSheet.Range["Q10"].Value = "Financiamiento";
                        xlWorkSheet.Range["R10"].Value = "APP";
                        xlWorkSheet.Range["S10"].Value = "Crédito";
                        xlWorkSheet.Range["T10"].Value = "Doc. SAP";

                        int lineabase = 11;
                        int lineatope = lineabase + ViasPago.Count;
                        int lineatope2 = lineabase + ViasPago.Count + 2;
                        int lineatope3 = lineabase + ViasPago.Count + 3;
                        int lineatope4 = lineabase + ViasPago.Count + 4;
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
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO);
                                    xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO);
                                    xlWorkSheet.Range["F" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }

                                MONTO = MONTO + Convert.ToDouble(ViasPago[k].MONTO);

                                xlWorkSheet.Range["G" + Convert.ToString(linea)].Value =ViasPago[k].NAME1;

                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["H" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                if (ViasPago[k].MONEDA == "CLP")
                                {
                                    MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                }
                                xlWorkSheet.Range["I" + Convert.ToString(linea)].Value = ViasPago[k].MONEDA; 
                                xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = ViasPago[k].NUM_CHEQUE;
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_DIA);
                                    xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                    xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_FECHA);
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_TRANSF);
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_VALE_V);
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_DEP);
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_TARJ);
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_FINANC);
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_APP);
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_CREDITO);
                                    xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                    xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                xlWorkSheet.Range["T" + Convert.ToString(linea)].Value = ViasPago[k].DOC_SAP;

                                // LLENAS EFECTIVO OTRAS MONEDAS
                                if (ViasPago[k].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(ViasPago[k].MONTO_EFEC));
                                    MONTO4 = MONTO4 + Convert.ToDouble(MonedaFormateada);
                                    xlWorkSheet.Range["E" + Convert.ToString(lineatope2)].Value = MONTO4;
                                }
                                if (ViasPago[k].MONEDA == "USD")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(ViasPago[k].MONTO_EFEC));
                                    MONTO2 = MONTO2 + Convert.ToDouble(MonedaFormateada);
                                    xlWorkSheet.Range["E" + Convert.ToString(lineatope3)].Value = MONTO2;
                                }
                                if (ViasPago[k].MONEDA == "EUR")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(ViasPago[k].MONTO_EFEC));
                                    MONTO3 = MONTO3 + Convert.ToDouble(MonedaFormateada);
                                    xlWorkSheet.Range["E" + Convert.ToString(lineatope4)].Value = MONTO3;
                                }
                            }
                        }
                        //TOTALES
                          xlWorkSheet.Range["D" + Convert.ToString(lineatope)].Value = "Totales: ";
                          xlWorkSheet.Range["D" + Convert.ToString(lineatope2)].Value = "Totales CLP: ";
                          xlWorkSheet.Range["D" + Convert.ToString(lineatope3)].Value = "Totales USD: ";
                          xlWorkSheet.Range["D" + Convert.ToString(lineatope4)].Value = "Totales EUR: ";
                        //Monto
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO));
                            xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = MonedaFormateada;
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_EFEC));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_DIA));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_FECHA));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_TRANSF));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_VALE_V));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_DEP));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_TARJ));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_FINANC));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_APP));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["S" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["S" + Convert.ToString(lineatope)].Value = MonedaFormateada;
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
                        xlWorkSheet.Range["K10"].Value = "Moneda";
                        xlWorkSheet.Range["L10"].Value = "Chq. al día";
                        xlWorkSheet.Range["M10"].Value = "Chq. a fecha";
                        xlWorkSheet.Range["N10"].Value = "Transferencia";
                        xlWorkSheet.Range["O10"].Value = "V. Vista";
                        xlWorkSheet.Range["P10"].Value = "Depósitos";
                        xlWorkSheet.Range["Q10"].Value = "Tarjetas";
                        xlWorkSheet.Range["R10"].Value = "Financiamiento";
                        xlWorkSheet.Range["S10"].Value = "APP";
                        xlWorkSheet.Range["T10"].Value = "Crédito";
                        xlWorkSheet.Range["U10"].Value = "Total Cajero";

                        int lineabase = 11;
                        int lineatope = lineabase + ViasPago.Count;
                        int lineatope2 = lineabase + ViasPago.Count + 2;
                        int lineatope3 = lineabase + ViasPago.Count + 3;
                        int lineatope4 = lineabase + ViasPago.Count + 4;

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
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].TOTAL_MOV);
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
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].TOTAL_INGR);
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
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["J" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                //MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                xlWorkSheet.Range["K" + Convert.ToString(linea)].Value = ViasPago[k].MONEDA; 
                                //Cheque al día
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_DIA);
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DIA);
                                    xlWorkSheet.Range["L" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DIA = MONTO_DIA + Convert.ToDouble(ViasPago[k].MONTO_DIA);
                                //Cheque a fecha
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_FECHA);
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FECHA);
                                    xlWorkSheet.Range["M" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FECHA = MONTO_FECHA + Convert.ToDouble(ViasPago[k].MONTO_FECHA);
                                //Transferencias
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_TRANSF);
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TRANSF);
                                    xlWorkSheet.Range["N" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TRANSF = MONTO_TRANSF + Convert.ToDouble(ViasPago[k].MONTO_TRANSF);
                                //Vale vistas
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_VALE_V);
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_VALE_V);
                                    xlWorkSheet.Range["O" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_VALE_V = MONTO_VALE_V + Convert.ToDouble(ViasPago[k].MONTO_VALE_V);
                                //Depositos
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_DEP);
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_DEP);
                                    xlWorkSheet.Range["P" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_DEP = MONTO_DEP + Convert.ToDouble(ViasPago[k].MONTO_DEP);
                                //Tarjetas
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_TARJ);
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_TARJ);
                                    xlWorkSheet.Range["Q" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_TARJ = MONTO_TARJ + Convert.ToDouble(ViasPago[k].MONTO_TARJ);
                                //Financiamiento
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_FINANC);
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_FINANC);
                                    xlWorkSheet.Range["R" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_FINANC = MONTO_FINANC + Convert.ToDouble(ViasPago[k].MONTO_FINANC);
                                //App
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_APP);
                                    xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_APP);
                                    xlWorkSheet.Range["S" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_APP = MONTO_APP + Convert.ToDouble(ViasPago[k].MONTO_APP);
                                //Credito
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_CREDITO);
                                    xlWorkSheet.Range["T" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_CREDITO);
                                    xlWorkSheet.Range["T" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                MONTO_CREDITO = MONTO_CREDITO + Convert.ToDouble(ViasPago[k].MONTO_CREDITO);
                                //Total Cajero
                                if (logApertura2[0].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].TOTAL_CAJERO);
                                    xlWorkSheet.Range["U" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                else
                                {
                                    MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].TOTAL_CAJERO);
                                    xlWorkSheet.Range["U" + Convert.ToString(linea)].Value = MonedaFormateada;
                                }
                                TOTAL_CAJERO = TOTAL_CAJERO + Convert.ToDouble(ViasPago[k].TOTAL_CAJERO);

                                // LLENAS EFECTIVO OTRAS MONEDAS
                                if (ViasPago[k].MONEDA == "CLP")
                                {
                                    MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(ViasPago[k].MONTO_EFEC));
                                    MONTO4 = MONTO4 + Convert.ToDouble(MonedaFormateada);
                                    xlWorkSheet.Range["E" + Convert.ToString(lineatope2)].Value = MONTO4;
                                }
                                if (ViasPago[k].MONEDA == "USD")
                                {
                                    MONTO2 = MONTO2 + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["E" + Convert.ToString(lineatope3)].Value = MONTO2;
                                }
                                if (ViasPago[k].MONEDA == "EUR")
                                {
                                    MONTO3 = MONTO3 + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                                    xlWorkSheet.Range["E" + Convert.ToString(lineatope4)].Value = MONTO3;
                                }
                            }
                        }

                        //TOTALES
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope)].Value = "Totales: ";
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope2)].Value = "Totales CLP: ";
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope3)].Value = "Totales USD: ";
                        xlWorkSheet.Range["D" + Convert.ToString(lineatope4)].Value = "Totales EUR: ";

                        //TOTALES
                        xlWorkSheet.Range["F" + Convert.ToString(lineatope)].Value = "Totales: ";
                        //Total movimientos
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(TOTAL_MOV));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(TOTAL_INGR));
                            xlWorkSheet.Range["I" + Convert.ToString(lineatope)].Value = Convert.ToString(MonedaFormateada);
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_INGR));
                            xlWorkSheet.Range["I" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Efectivo
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO4));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_DIA));
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DIA));
                            xlWorkSheet.Range["L" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Cheque a fecha
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_FECHA));
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FECHA));
                            xlWorkSheet.Range["M" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Transferencia
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_TRANSF));
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TRANSF));
                            xlWorkSheet.Range["N" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Vale vista
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_VALE_V));
                            xlWorkSheet.Range["O" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_VALE_V));
                            xlWorkSheet.Range["O" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Deposito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_DEP));
                            xlWorkSheet.Range["P" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_DEP));
                            xlWorkSheet.Range["P" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Tarjetas
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_TARJ));
                            xlWorkSheet.Range["Q" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_TARJ));
                            xlWorkSheet.Range["Q" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Financiamiento
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_FINANC));
                            xlWorkSheet.Range["R" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_FINANC));
                            xlWorkSheet.Range["R" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //APP
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_APP));
                            xlWorkSheet.Range["S" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_APP));
                            xlWorkSheet.Range["S" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Credito
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["T" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(MONTO_CREDITO));
                            xlWorkSheet.Range["T" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        //Total cajero
                        if (logApertura2[0].MONEDA == "CLP")
                        {
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(TOTAL_CAJERO));
                            xlWorkSheet.Range["U" + Convert.ToString(lineatope)].Value = MonedaFormateada;
                        }
                        else
                        {
                            MonedaFormateada = FM.FormatoMonedaExtranjera(Convert.ToString(TOTAL_CAJERO));
                            xlWorkSheet.Range["U" + Convert.ToString(lineatope)].Value = MonedaFormateada;
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_EFEC);
                                xlWorkSheet.Range["E" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            else
                            {
                                MonedaFormateada = FM.FormatoMonedaExtranjera(ViasPago[k].MONTO_EFEC);
                                xlWorkSheet.Range["E" + Convert.ToString(linea)].Value = MonedaFormateada;
                            }
                            //MONTO_EFEC = MONTO_EFEC + Convert.ToDouble(ViasPago[k].MONTO_EFEC);
                            //Cheque al dia
                            if (logApertura2[0].MONEDA == "CLP")
                            {
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_DIA);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_FECHA);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_TRANSF);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_VALE_V);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_DEP);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_TARJ);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_FINANC);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_APP);
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
                                MonedaFormateada = FM.FormatoMoneda2(ViasPago[k].MONTO_CREDITO);
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_EFEC));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_DIA));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_FECHA));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_TRANSF));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_VALE_V));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_DEP));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_TARJ));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_FINANC));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_APP));
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
                            MonedaFormateada = FM.FormatoMoneda2(Convert.ToString(MONTO_CREDITO));
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

                System.Windows.Forms.MessageBox.Show("Archivo Excel creado");
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
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
