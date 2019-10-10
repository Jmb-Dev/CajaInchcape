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
using CajaIndigo.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;
using CajaIndigo.AppPersistencia.Class.AutorizadorAnulaciones;

namespace CajaIndigo
{
    /// <summary>
    /// Interaction logic for Autorizacion.xaml
    /// </summary>
    public partial class Autorizacion : Window
    {
        public Autorizacion(string User, string Password, string IdCaja)
        {
            InitializeComponent();

            txtUser.Text = User;
            txtPassword.Text = Password;
            //txtUserAutoriza.Text = UserAutoriza;
            //txtPass.Text = Pass;
            txtIdCaja2.Text = IdCaja;
            GC.Collect();
        }

        private void btnAutoriza_Click(object sender, RoutedEventArgs e)
        {
            AutorizaAnulaciones autorizaanulaciones = new AutorizaAnulaciones();
            //RFC AUTORIZACION DE ANULACIONES DE COMPROBANTES
            txtUserAutoriza.Text = txtUserAutoriza.Text.ToUpper();

            autorizaanulaciones.anulacioncomprobantes(txtUser.Text, txtPassword.Text, txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtUserAutoriza.Text, txtPass.Password, txtIdCaja2.Text);

            //AutorizacionInterfaz formInterface = this.Owner as AutorizacionInterfaz;
            //if (formInterface != null)
            //{
            //    formInterface.BuscaAutorizacion(autorizaanulaciones.Autorizado);

            //}

            //PagosDocumentos pagosdocumentos = PagosDocumentos.GetWindow(this);
            //Window pagosdocumentos = Window.GetWindow(this);
            //pagosdocumentos.GBDetalleDocs.Visibility = Visibility.Visible;
            Vista.PagoDocumento.PagoDocumento window = new Vista.PagoDocumento.PagoDocumento();

           // PagosDocumentos window = Window.GetWindow(this.Owner) as PagosDocumentos;
            if(window != null)
            {
                this.Close();
                if (autorizaanulaciones.Valido == true)
                {
                
                //window.ListaDocumentosAnulacion();
               // window.txtUserAnula.Text = txtUserAutoriza.Text;
              //  window.txtComentAnula.Text = txtComment.Text;
                //window.btnAnular.IsEnabled = true;
               
                }
                
            }

            GC.Collect();
        }
    }
}
