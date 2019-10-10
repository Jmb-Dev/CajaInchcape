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
using CajaIndigo.AppPersistencia.Class.UsuariosCaja;
using CajaIndigo.AppPersistencia.Class.BloquearCaja;
using CajaIndigo;
using System.Windows.Threading;

namespace CajaIndigo.Vista.Menu
{
    /// <summary>
    /// Interaction logic for MenuCaja.xaml
    /// </summary>
    public partial class MenuCaja : System.Windows.Window
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

        
        public MenuCaja(string usuariologg, string passlogg, string usuariotemp, string cajaconect, string sucursal, string sociedad, List<string> moneda, string pais, double monto, List<LOG_APERTURA> logApertura)
            {
                try
                {
                   InitializeComponent();

                    if (moneda.Count == 0)
                    {
                        moneda.Add("CLP");
                        moneda.Add("USD");
                    }
                    int test = 0;

                    GBInicio.Visibility = Visibility.Visible;
                    textBlock6.Content = cajaconect;
                    textBlock7.Content = usuariologg;
                    textBlock8.Content = sucursal;
                    textBlock9.Content = usuariotemp ;
                    lblMonto.Content = Convert.ToString(monto);
                    lblSociedad.Content = sociedad;
                    
                    lblPassword.Content = passlogg;
                    DateTime result = DateTime.Today;
                    calendario.Text = Convert.ToString(result);
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

              private void BloquearCaja_Click(object sender, RoutedEventArgs e)
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
                  bloquearcaja.bloqueardesbloquearcaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, logApertura2);
                  this.Close();
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

              private void BloquearCaja_Click_1(object sender, RoutedEventArgs e)
              {
                  bloquearCaja();
              }

              private void bloquearCaja()
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
                  bloquearcaja.bloqueardesbloquearcaja(Convert.ToString(textBlock7.Content), Convert.ToString(lblPassword.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, logApertura2);
                  this.Close();
                  GC.Collect();

              }
    }
}
