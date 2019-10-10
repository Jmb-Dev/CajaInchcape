using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Web.UI.WebControls;
using iTextSharp.text;
using System.Drawing.Imaging;
using iTextSharp.text.pdf;
using System.Collections;
using System.Windows.Controls.Primitives;
using CajaIndu.AppPersistencia.Class.BusquedaReimpresiones.Estructura;
using CajaIndu.AppPersistencia.Class.BusquedaReimpresiones;
using CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura;



namespace CajaIndu
{
    /// <summary>
    /// Interaction logic for Comprobante.xaml
    /// </summary>
    public partial class Comprobante : Window
    {
        List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP> ListViasPagosAux = new List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP>();
        List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS> DocsAPagarAux = new List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS>();
        System.Drawing.Image bitmap;

        public Comprobante(List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP> ListViasPagos
            , List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS> DocsAPagar
            , string NomCliente, string RUTCliente, string Cajero, string Usuario, string NomCaja, string Ingreso, string NotaVenta
            , string DocContable, string InOut, string Moneda, string Pedido, string Mandante)
        {
            InitializeComponent();

            ListViasPagosAux.Clear();
            DocsAPagarAux.Clear();
            ListViasPagosAux = ListViasPagos;
            DocsAPagarAux = DocsAPagar;
            //this.Visibility = Visibility.Collapsed;
            DateTime result = DateTime.Today;
            txtFecha.Text= Convert.ToString(result.Date).Substring(0,10);
            txtHora.Text = Convert.ToString(result.TimeOfDay);//.Substring(11);
            DGPagos.ItemsSource = DocsAPagar;
            DGResumenViasPago.ItemsSource = ListViasPagos;
            txtNomCli.Text = NomCliente;
            txtRUT.Text =RUTCliente;
            txtCajero.Text = Cajero;
            txtUsuario.Text = Usuario;
            txtNomCaja.Text = NomCaja;
            txtIngreso.Text = Convert.ToString(Ingreso);
            txtNotaVta.Text = NotaVenta;
            txtNumDocCont.Text = DocContable;
            txtPagina.Text = "";
            txtMoneda.Text = Moneda;
            txtInOut.Text = InOut;
            txtPedido.Text = Pedido;
            txtMandante.Text = Mandante;
            //GeneracionDePDF(ListViasPagosAux, DocsAPagarAux);
           // this.Close();
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

        private void GeneracionDePDF_Nuevo(List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP> ListViasPagosAux, List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS> DocsAPagarAux, string InOut)
        {
            try
            {
                string appRootDir = Convert.ToString(System.IO.Path.GetTempPath());
                //string appRootDir = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName;
                string startFile = appRootDir + "/PDFs/" + "Chapter1_Example5.pdf";
                string watermarkedFile = appRootDir + "InduLog\\" + txtIngreso.Text + "-Nuevo.Text.pdf";
			    string unwatermarkedFile = appRootDir + "/PDFs/" + "Chapter1_Example5_Un-Watermarked.pdf";
                string direct = Convert.ToString(System.IO.Path.GetTempPath());
                direct = direct + "InduLog\\" + txtIngreso.Text + ".pdf";
               // direct = direct + "InduLog\\ResumenMensualMovimientos29092014.pdf";

			string watermarkText = "No válido como comprobante";
            //Document pdfcommande = new Document(PageSize.LETTER);
			// Creating a Five paged PDF
            using (FileStream fs = new FileStream(direct, FileMode.Create, FileAccess.Write, FileShare.None))
               
            using (Document pdfcommande = new Document(PageSize.LETTER,20f,20f,100f,100f))
			//using (Document doc = new Document(PageSize.LETTER))
			using (PdfWriter writer = PdfWriter.GetInstance(pdfcommande, fs))
            {
               // Document pdfcommande = new Document(PageSize.LETTER);
               
                //PdfReader reader = new PdfReader(Convert.ToString(System.IO.Path.GetTempPath()) + "InduLog\\" + "RendicionCaja20102014" + ".pdf");
                txtDirect.Text = direct;
                try
                {
                    //***PdfWriter.GetInstance(pdfcommande, new FileStream(direct, FileMode.OpenOrCreate));
                    //PdfReader reader = new PdfReader(direct+"reader");
                    pdfcommande.Open();

                    pdfcommande.NewPage();

                    
                   // // add image to document
                   // string DirectPhotos = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName;
                   ////\\sap-pos-des.ad.gi.cl\Proyecto.Cajas.DBM\Imagenes Caja
                   //// string DirectImages = System.IO.Directory.GetCurrentDirectory();
                   // string DirectImages = "\\\\sap-pos-des.ad.gi.cl\\Proyecto.Cajas.DBM\\Imagenes Caja";
                   // //string DirectImages = Convert.ToString(System.IO.Path.GetTempPath());
                   // //string PruebaDirect = DirectImages + "\\logo_empresas_indumotora_sociedad_EI15.png";
                   // iTextSharp.text.Image logo; 
                   
                   // if (txtSociedad.Text == "EI15")
                   // {
                   //     //DirectImages = DirectImages + CajaIndu.Properties.Resources.SociedadEI15.ToString();
                   //     DirectImages = DirectImages + "\\SociedadEI15.png";
                   //     logo = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.SociedadEI15, System.Drawing.Imaging.ImageFormat.Bmp);
                   //     logo.Alignment = iTextSharp.text.Image.ALIGN_LEFT;
                   //     logo.ScalePercent(25f);
                   //     pdfcommande.Add(logo);
                   // }
                   // if (txtSociedad.Text == "EI17")
                   // {
                   //     //logo = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.SociedadEI17);
                   //     //var ei17 = new Bitmap(CajaIndu.Properties.Resources.SociedadEI17);
                   //     //DirectImages = DirectImages + CajaIndu.Properties.Resources.SociedadEI17.ToString();
                   //     DirectImages = DirectImages + "\\SociedadEI17.png";
                   //     logo = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.SociedadEI17, System.Drawing.Imaging.ImageFormat.Bmp);
                   //     logo.Alignment = iTextSharp.text.Image.ALIGN_LEFT;
                   //     logo.ScalePercent(25f);
                   //     pdfcommande.Add(logo);
                   // }
                    //iTextSharp.text.Image logo = iTextSharp.text.Image.GetInstance(DirectImages);
                  
                   

                    //for (int page = 1; page <= reader.NumberOfPages; page++)
                    //{

                    //}

                    ////



                    // string texto = "Este es mi primer PDF al vuelo";
                    // pdfcommande.Add(new iTextSharp.text.Paragraph(texto));
                    // pdfcommande.Close();


                    //  Document pdfcommande = new Document();
                    PdfPTable table = new PdfPTable(DGPagos.Columns.Count);
                    table.TotalWidth = 580f;
                    table.LockedWidth = true;
                    table.SpacingBefore = 20f;
                    table.SpacingAfter = 30f;
                    // table.WidthPercentage = 100;

                    List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS> Docs = new List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS>();

                    for (int k = 0; k < DGPagos.Items.Count; k++)
                    {

                        if (k == 0)
                        {
                            DGPagos.Items.MoveCurrentToFirst();
                        }
                        Docs.Add(DGPagos.Items.CurrentItem as CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_DOCUMENTOS);
                        DGPagos.Items.MoveCurrentToNext();
                    }
                    PdfPCell cell = new PdfPCell(new Phrase(Convert.ToString(label12.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    //PdfPCell cell2 = new PdfPCell();
                    cell.Colspan = DGPagos.Columns.Count;
                    cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right

                    table.AddCell(cell);
                    //DataGridColumn Colun = new DataGridColumn();

                    //in DGPagos.Columns
                    foreach (DataGridColumn column in DGPagos.Columns)
                    {
                        table.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    }
                    table.HeaderRows = 1;

                    string FechaAFormatear;
                    string Dia;
                    string Mes;
                    string Ano;
                    FormatoMonedas FM = new FormatoMonedas();
                    string MonedaFormateada;
                    for (int k = 0; k < Docs.Count; k++)
                    {
                        if (Docs[k] != null)
                        {
                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(Docs[k].TXT_DOCU), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table.AddCell(cellrow1);
                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(Docs[k].NRO_DOCUMENTO), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table.AddCell(cellrow2);
                            Dia = Docs[k].FECHA_DOC.Substring(8, 2);
                            Mes = Docs[k].FECHA_DOC.Substring(5, 2);
                            Ano = Docs[k].FECHA_DOC.Substring(0, 4);
                            FechaAFormatear = Dia + "/" + Mes + "/" + Ano;
                            PdfPCell cellrow3 = new PdfPCell(new Phrase(String.Format("{0:dd/MM/yyyy}", FechaAFormatear), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table.AddCell(cellrow3);
                            Dia = Docs[k].FECHA_VENC_DOC.Substring(8, 2);
                            Mes = Docs[k].FECHA_VENC_DOC.Substring(5, 2);
                            Ano = Docs[k].FECHA_VENC_DOC.Substring(0, 4);
                            FechaAFormatear = Dia + "/" + Mes + "/" + Ano;

                            PdfPCell cellrow4 = new PdfPCell(new Phrase(String.Format("{0:dd/MM/yyyy}", FechaAFormatear), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table.AddCell(cellrow4);
                            if (txtMoneda.Text == "CLP")
                                MonedaFormateada = FM.FormatoMonedaCaja(Docs[k].MONTO_DOC_ML, "Ch", "1");
                            else
                                MonedaFormateada = FM.FormatoMonedaCaja(Docs[k].MONTO_DOC_ML, "Ex", "1");
                            PdfPCell cellrow5 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table.AddCell(cellrow5);
                            if (txtMoneda.Text == "CLP")
                                MonedaFormateada = FM.FormatoMonedaCaja(Docs[k].MONTO_DOC_MO, "Ch", "1");
                            else
                                MonedaFormateada = FM.FormatoMonedaCaja(Docs[k].MONTO_DOC_MO, "Ex", "1");
                            PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table.AddCell(cellrow6);

                        }
                    }
                    IEnumerable itemsSource = DGPagos.ItemsSource as IEnumerable;

                    PdfPTable table2 = new PdfPTable(DGResumenViasPago.Columns.Count);
                    table2.TotalWidth = 580f;
                    table2.LockedWidth = true;
                    table2.SpacingBefore = 20f;
                    table2.SpacingAfter = 30f;
                   
                    List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP> ViasPago = new List<CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP>();
                    for (int k = 0; k < DGResumenViasPago.Items.Count; k++)
                    {

                        if (k == 0)
                        {
                            DGResumenViasPago.Items.MoveCurrentToFirst();
                        }
                        ViasPago.Add(DGResumenViasPago.Items.CurrentItem as CajaIndu.AppPersistencia.Class.ReimpresionComprobantes.Estructura.DATOS_VP);
                        DGResumenViasPago.Items.MoveCurrentToNext();
                    }
                    PdfPCell cell2 = new PdfPCell(new Phrase(Convert.ToString(label11.Content), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    //PdfPCell cell2 = new PdfPCell();
                    cell2.Colspan = DGResumenViasPago.Columns.Count;
                    cell2.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                    table2.AddCell(cell2);
                    foreach (DataGridColumn column in DGResumenViasPago.Columns)
                    {
                        table2.AddCell(new Phrase(column.Header.ToString(), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));

                    }
                    table2.HeaderRows = 1;
                    for (int k = 0; k < ViasPago.Count; k++)
                    {
                        if (ViasPago[k] != null)
                        {
                            PdfPCell cellrow1 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NUM_POS), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow1);
                            PdfPCell cellrow2 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].DESCRIP_VP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow2);
                            PdfPCell cellrow3 = new PdfPCell(new Phrase(Convert.ToString(ViasPago[k].NUM_VP), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow3);
                            Dia = ViasPago[k].FECHA_EMISION.Substring(8, 2);
                            Mes = ViasPago[k].FECHA_EMISION.Substring(5, 2);
                            Ano = ViasPago[k].FECHA_EMISION.Substring(0, 4);
                            FechaAFormatear = Dia + "/" + Mes + "/" + Ano;
                            PdfPCell cellrow4 = new PdfPCell(new Phrase(String.Format("{0:dd/MM/yyyy}", FechaAFormatear), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow4);
                            Dia = ViasPago[k].FECHA_VENC.Substring(8, 2);
                            Mes = ViasPago[k].FECHA_VENC.Substring(5, 2);
                            Ano = ViasPago[k].FECHA_VENC.Substring(0, 4);
                            FechaAFormatear = Dia + "/" + Mes + "/" + Ano;
                            PdfPCell cellrow5 = new PdfPCell(new Phrase(String.Format("{0:dd/MM/yyyy}", FechaAFormatear), new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow5);
                            if (txtMoneda.Text == "CLP")
                                MonedaFormateada = FM.FormatoMonedaCaja(ViasPago[k].MONTO_ML, "Ch", "1");
                            else
                                MonedaFormateada = FM.FormatoMonedaCaja(ViasPago[k].MONTO_ML, "Ex", "1");
                            PdfPCell cellrow6 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow6);
                            if (txtMoneda.Text == "CLP")
                                MonedaFormateada = FM.FormatoMonedaCaja(ViasPago[k].MONTO_MO, "Ch", "1");
                            else
                                MonedaFormateada = FM.FormatoMonedaCaja(ViasPago[k].MONTO_MO, "Ex", "1");
                            PdfPCell cellrow7 = new PdfPCell(new Phrase(MonedaFormateada, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 8f, iTextSharp.text.Font.NORMAL)));
                            table2.AddCell(cellrow7);
                        }
                    }

                    pdfcommande.Add(iTextSharp.text.PageSize.LETTER);
                   
                    string texto = "";
                    //Titulo
                    texto = "Comprobante de " + InOut;
                    iTextSharp.text.Paragraph itxtTitulo = new iTextSharp.text.Paragraph(texto);
                    itxtTitulo.IndentationLeft = 200;
                    itxtTitulo.Font.Size = 10;
                    itxtTitulo.Font.SetStyle("bold");
                    itxtTitulo.Font.SetFamily("courier");
                    //itxtTitulo.SpacingBefore = 5f;
                    itxtTitulo.SpacingAfter = 5f;
                    pdfcommande.Add(itxtTitulo);
                    //DATOS CAJA
                    texto = Convert.ToString(label6.Content) + txtNomCaja.Text + "     " + Convert.ToString(label7.Content) + txtCajero.Text;
                    iTextSharp.text.Paragraph itxtHeader = new iTextSharp.text.Paragraph(texto);
                    itxtHeader.IndentationLeft = 10;
                    itxtHeader.Font.Size = 9;
                    itxtHeader.Font.SetFamily("courier");
                    pdfcommande.Add(itxtHeader);
                    texto = "  ";
                    pdfcommande.Add(new iTextSharp.text.Paragraph(texto));
                    //DATOS CLIENTE
                    //TABLE F
                    PdfPTable tablef = new PdfPTable(4);
                    tablef.TotalWidth = 400f;
                    tablef.LockedWidth = true;
                    tablef.HorizontalAlignment = 0;
                    float[] widths = new float[] { 40f, 40f, 50f, 120f};
                    tablef.SetWidths(widths);
                    //cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                    PdfPCell cellrow1f = new PdfPCell(new Phrase("RUT Cliente:", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    //cellrow1f.Border = 8;
                    cellrow1f.Left = 0;
                    cellrow1f.HorizontalAlignment = 0;
                    tablef.AddCell(cellrow1f); //, iTextSharp.text.Font.NORMAL,iTextSharp.text.BaseColor.WHITE
                    PdfPCell cellrow2f = new PdfPCell(new Phrase(txtRUT.Text, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    //cellrow2f.Border = 8;
                    cellrow2f.HorizontalAlignment = 0;
                    cellrow2f.Left = 200f;
                    tablef.AddCell(cellrow2f);
                    PdfPCell cellrow3f = new PdfPCell(new Phrase("Nombre cliente:", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    //cellrow3f.Border = 8;
                    cellrow3f.Left = 160f;
                    cellrow3f.HorizontalAlignment = 0;
                    tablef.AddCell(cellrow3f);
                    PdfPCell cellrow4f = new PdfPCell(new Phrase(txtNomCli.Text, new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.COURIER, 9f, iTextSharp.text.Font.NORMAL)));
                    //cellrow4f.Border = 8;
                    cellrow4f.HorizontalAlignment = 0;
                    tablef.AddCell(cellrow4f);
                    pdfcommande.Add(tablef);
                    //Datos de cliente
                    //texto = txtRUT.Text + "    " + txtNomCli.Text;
                    //iTextSharp.text.Paragraph itxtCliente = new iTextSharp.text.Paragraph(texto);
                    //itxtCliente.IndentationLeft = 5;
                    //itxtCliente.Font.Size = 9;
                    //itxtCliente.Font.SetFamily("courier");
                    //pdfcommande.Add(itxtCliente);
                    
                    //TABLAS DEL DETALLES DEL COMPROBANTE
                    //pdfcommande.Add(new iTextSharp.text.Paragraph(Convert.ToString(label12.Content)));
                    pdfcommande.Add(table);
                    //pdfcommande.Add(new iTextSharp.text.Paragraph(Convert.ToString(label11.Content)));
                    pdfcommande.Add(table2);
                    // pdfcommande.Add(firstpara);


                    //DirectImages = System.IO.Directory.GetCurrentDirectory();
                    //DirectImages = DirectImages + "\\logos.png";
                    //iTextSharp.text.Image logos = iTextSharp.text.Image.GetInstance(DirectImages);
                    //logos.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                    //logos.ScalePercent(25f);
                    //logos.Bottom = 20f;
                    //pdfcommande.Add(logos);

                    //DirectImages = System.IO.Directory.GetCurrentDirectory();
                    //DirectImages = DirectImages + "\\barra-contacto.png";
                    //iTextSharp.text.Image barra = iTextSharp.text.Image.GetInstance(DirectImages);
                    //barra.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                    //barra.ScalePercent(25f);
                    //barra.Bottom = 10f;
                    //pdfcommande.Add(barra);
                }


                catch (Exception ex)
                {
                    Console.Write(ex.Message, ex.StackTrace);
                  //  System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                }


                    pdfcommande.Close();
                    pdfcommande.Dispose();
                }
                   // pdfcommande.Dispose();

            try
            {
                // Creating watermark on a separate layer
                // Creating iTextSharp.text.pdf.PdfReader object to read the Existing PDF Document produced by 1 no.
                string direct2 = Convert.ToString(System.IO.Path.GetTempPath());

                direct2 = direct2 + "InduLog\\ResumenMensualMovimientos29092014.pdf";
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
                        if ((txtMandante.Text == "100") | (txtMandante.Text == "200"))
                        {
                            // Get the ContentByte object
                            PdfContentByte cb = stamper.GetUnderContent(i);
                            // Tell the cb that the next commands should be "bound" to this new layer
                            cb.BeginLayer(layer);
                            cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 50);
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
                        cb2.ShowTextAligned(PdfContentByte.ALIGN_LEFT, Paginas, rect.Width - 170, rect.Height - 70, 0f);
                        cb2.EndText();
                        // Close the layer
                        cb2.EndLayer();

                        //EMPRESA
                        //PdfContentByte cb3 = stamper.GetUnderContent(i);
                        //cb3.BeginLayer(layer);
                        //cb3.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        //gState2.FillOpacity = 1f;
                        //cb3.SetGState(gState2);
                        //cb3.SetColorFill(BaseColor.BLACK);
                        //cb3.BeginText();
                        //cb3.ShowTextAligned(PdfContentByte.ALIGN_LEFT, txtEmpresa.Text,20, rect.Height - 100, 0f);
                        //cb3.EndText();
                        //// Close the layer
                        //cb3.EndLayer();
                        //FECHA LABEL
                        PdfContentByte cb4a = stamper.GetUnderContent(i);
                        cb4a.BeginLayer(layer);
                        cb4a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb4a.SetGState(gState2);
                        cb4a.SetColorFill(BaseColor.BLACK);
                        cb4a.BeginText();
                        cb4a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Fecha:", rect.Width - 170, rect.Height - 80, 0f);
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
                        cb4.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, String.Format("{0:dd/MM/yyyy}", FechaAFormatear), rect.Width - 30, rect.Height - 80, 0f);
                        cb4.EndText();
                        // Close the layer
                        cb4.EndLayer();

                        // RUT EMPRESA
                        //PdfContentByte cb5 = stamper.GetUnderContent(i);
                        //cb5.BeginLayer(layer);
                        //cb5.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        //gState2.FillOpacity = 1f;
                        //cb5.SetGState(gState2);
                        //cb5.SetColorFill(BaseColor.BLACK);
                        //cb5.BeginText();
                        //cb5.ShowTextAligned(PdfContentByte.ALIGN_LEFT, txtRIF.Text  , 20, rect.Height - 80, 0f);
                        //cb5.EndText();
                        //// Close the layer
                        //cb5.EndLayer();

                        //HORA
                        PdfContentByte cb6 = stamper.GetUnderContent(i);
                        cb6.BeginLayer(layer);
                        cb6.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb6.SetGState(gState2);
                        cb6.SetColorFill(BaseColor.BLACK);
                        cb6.BeginText();
                        FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                        cb6.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, String.Format("{0:HH:mm:ss}", FechaAFormatear), rect.Width - 30, rect.Height - 90, 0f);
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
                        cb6a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Hora:", rect.Width - 170, rect.Height - 90, 0f);
                        cb6a.EndText();
                        // Close the layer
                        cb6a.EndLayer();

                        // SOCIEDAD
                        //PdfContentByte cb7 = stamper.GetUnderContent(i);
                        //cb7.BeginLayer(layer);
                        //cb7.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        //gState2.FillOpacity = 1f;
                        //cb7.SetGState(gState2);
                        //cb7.SetColorFill(BaseColor.BLACK);
                        //cb7.BeginText();
                        //cb7.ShowTextAligned(PdfContentByte.ALIGN_LEFT, txtSociedad.Text, 20, rect.Height - 90, 0f);
                        //cb7.EndText();
                        //// Close the layer
                        //cb7.EndLayer();

                        //USUARIO
                        PdfContentByte cb8 = stamper.GetUnderContent(i);
                        cb8.BeginLayer(layer);
                        cb8.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb8.SetGState(gState2);
                        cb8.SetColorFill(BaseColor.BLACK);
                        cb8.BeginText();
                        FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                        cb8.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, txtUsuario.Text, rect.Width - 30, rect.Height - 100, 0f);
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
                        cb8a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Usuario:", rect.Width - 170, rect.Height - 100, 0f);
                        cb8a.EndText();
                        // Close the layer
                        cb8a.EndLayer();

                        //INGRESO
                        PdfContentByte cb9 = stamper.GetUnderContent(i);
                        cb9.BeginLayer(layer);
                        cb9.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb9.SetGState(gState2);
                        cb9.SetColorFill(BaseColor.BLACK);
                        cb9.BeginText();
                        FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                        cb9.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, txtIngreso.Text, rect.Width - 30, rect.Height - 110, 0f);
                        cb9.EndText();
                        // Close the layer
                        cb9.EndLayer();

                        //INGRESO LABEL
                        PdfContentByte cb9a = stamper.GetUnderContent(i);
                        cb9a.BeginLayer(layer);
                        cb9a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb9a.SetGState(gState2);
                        cb9a.SetColorFill(BaseColor.BLACK);
                        cb9a.BeginText();
                        cb9a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Ingreso:", rect.Width - 170, rect.Height - 110, 0f);
                        cb9a.EndText();
                        // Close the layer
                        cb9a.EndLayer();

                        //NTA VENTA
                        PdfContentByte cb10 = stamper.GetUnderContent(i);
                        cb10.BeginLayer(layer);
                        cb10.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb10.SetGState(gState2);
                        cb10.SetColorFill(BaseColor.BLACK);
                        cb10.BeginText();
                        FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                        cb10.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, txtNotaVta.Text, rect.Width - 30, rect.Height - 120, 0f);
                        cb10.EndText();
                        // Close the layer
                        cb10.EndLayer();

                        //NTA VENTA LABEL
                        PdfContentByte cb10a = stamper.GetUnderContent(i);
                        cb10a.BeginLayer(layer);
                        cb10a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb10a.SetGState(gState2);
                        cb10a.SetColorFill(BaseColor.BLACK);
                        cb10a.BeginText();
                        cb10a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Nota de venta:", rect.Width - 170, rect.Height - 120, 0f);
                        cb10a.EndText();
                        // Close the layer
                        cb10a.EndLayer();

                        //DOC CONTABLE
                        PdfContentByte cb11 = stamper.GetUnderContent(i);
                        cb11.BeginLayer(layer);
                        cb11.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb11.SetGState(gState2);
                        cb11.SetColorFill(BaseColor.BLACK);
                        cb11.BeginText();
                        FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                        cb11.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, txtNumDocCont.Text, rect.Width - 30, rect.Height - 130, 0f);
                        cb11.EndText();
                        // Close the layer
                        cb11.EndLayer();

                        //DOC CONTABLE LABEL
                        PdfContentByte cb11a = stamper.GetUnderContent(i);
                        cb11a.BeginLayer(layer);
                        cb11a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb11a.SetGState(gState2);
                        cb11a.SetColorFill(BaseColor.BLACK);
                        cb11a.BeginText();
                        cb11a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "Doc. Contable:", rect.Width - 170, rect.Height - 130, 0f);
                        cb11a.EndText();
                        // Close the layer
                        cb11a.EndLayer();

                        //PEDIDO
                        PdfContentByte cb12 = stamper.GetUnderContent(i);
                        cb12.BeginLayer(layer);
                        cb12.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb12.SetGState(gState2);
                        cb12.SetColorFill(BaseColor.BLACK);
                        cb12.BeginText();
                        FechaAFormatear = Convert.ToString(DateTime.Now.Hour + ":" + Convert.ToString(DateTime.Now.Minute) + ":" + Convert.ToString(DateTime.Now.Second));
                        cb12.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, txtPedido.Text, rect.Width - 30, rect.Height - 140, 0f);
                        cb12.EndText();
                        // Close the layer
                        cb12.EndLayer();

                        //PEDIDO LABEL
                        PdfContentByte cb12a = stamper.GetUnderContent(i);
                        cb12a.BeginLayer(layer);
                        cb12a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                        gState2.FillOpacity = 1f;
                        cb12a.SetGState(gState2);
                        cb12a.SetColorFill(BaseColor.BLACK);
                        cb12a.BeginText();
                        cb12a.ShowTextAligned(PdfContentByte.ALIGN_LEFT, "N° Pedido:", rect.Width - 170, rect.Height - 140, 0f);
                        cb12a.EndText();
                        // Close the layer
                        cb12a.EndLayer();


                        if (txtSociedad.Text == "EI17")
                        {
                            if (i == 1)
                            {
                                PdfContentByte cb14a = stamper.GetUnderContent(i);
                                iTextSharp.text.Image Soc17 = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.logo_camiones1, System.Drawing.Imaging.ImageFormat.Jpeg);
                                Soc17.Alignment = iTextSharp.text.Image.ALIGN_LEFT;
                                cb14a.BeginLayer(layer);
                                cb14a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb14a.SetGState(gState2);
                                cb14a.SetColorFill(BaseColor.BLACK);
                                cb14a.BeginText();
                                Soc17.SetAbsolutePosition(20, 700);
                                Soc17.ScaleAbsolute(200, 40);
                                cb14a.AddImage(Soc17);   
                                cb14a.EndText();
                                cb14a.EndLayer();

                                PdfContentByte cb13a = stamper.GetUnderContent(i);
                                iTextSharp.text.Image logos = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.huincha_completa_camiones_y_buses1, System.Drawing.Imaging.ImageFormat.Jpeg);
                                logos.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                                logos.SetAbsolutePosition(10, rect.Height - (rect.Height - 70));
                                logos.ScaleAbsolute(600,75);    
                                cb13a.BeginLayer(layer);
                                cb13a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb13a.SetGState(gState2);
                                cb13a.SetColorFill(BaseColor.BLACK);
                                cb13a.BeginText();
                                cb13a.AddImage(logos);
                                cb13a.EndText();
                                // Close the layer
                                cb13a.EndLayer();
                            }
                        }

                        if (txtSociedad.Text == "EI15")
                        {
                            if (i == 1)
                            {
                                PdfContentByte cb14 = stamper.GetUnderContent(i);
                                //DirectImages = System.IO.Directory.GetCurrentDirectory();
                                iTextSharp.text.Image Soc15 = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.logo_indumotora_one1, System.Drawing.Imaging.ImageFormat.Jpeg);
                                Soc15.Alignment = iTextSharp.text.Image.ALIGN_LEFT;
                                Soc15.SetAbsolutePosition(20, 700);
                                Soc15.ScaleAbsolute(150, 50);
                                cb14.BeginLayer(layer);
                                cb14.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb14.SetGState(gState2);
                                cb14.SetColorFill(BaseColor.BLACK);
                                cb14.BeginText();
                                cb14.AddImage(Soc15);
                                cb14.EndText();
                                // Close the layer
                                cb14.EndLayer();

                                PdfContentByte cb13a = stamper.GetUnderContent(i);
                                iTextSharp.text.Image logos = iTextSharp.text.Image.GetInstance(CajaIndu.Properties.Resources.logos_one_png, System.Drawing.Imaging.ImageFormat.Png);
                                logos.Alignment = iTextSharp.text.Image.ALIGN_CENTER;
                                logos.SetAbsolutePosition(10, rect.Height - (rect.Height - 70));
                                logos.ScaleAbsolute(600, 70);                 
                                cb13a.BeginLayer(layer);
                                cb13a.SetFontAndSize(BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 9);
                                gState2.FillOpacity = 1f;
                                cb13a.SetGState(gState2);
                                cb13a.SetColorFill(BaseColor.BLACK);
                                cb13a.BeginText();
                                cb13a.AddImage(logos);
                                cb13a.EndText();
                             // Close the layer
                                cb13a.EndLayer();
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message, ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
              
                string url_reimpresion = "";
                url_reimpresion = watermarkedFile;
                PDFViewer pdfvisor = new PDFViewer();
                pdfvisor.webBrowser1.Navigate(url_reimpresion);
                pdfvisor.txtArchivo.Text = watermarkedFile;
                pdfvisor.txtArchivoNuevo.Text = direct;
                pdfvisor.Owner = this;
                pdfvisor.Show();
                this.Visibility = Visibility.Collapsed;
                GC.Collect();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                Console.Write(ex.Message, ex.StackTrace);
            }

        }
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            
        }
        
        private void Window_Closed(object sender, EventArgs e)
        {
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GeneracionDePDF_Nuevo(ListViasPagosAux, DocsAPagarAux, txtInOut.Text);
        }

        private void txtSociedad_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        
    }
}
