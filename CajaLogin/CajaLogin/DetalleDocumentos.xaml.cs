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
//using CajaIndu.AppPersistencia.Class.PartidasAbiertas.Estructura;
//using CajaIndu.AppPersistencia.Class.PartidasAbiertas;

namespace CajaIndu
{
    /// <summary>
    /// Interaction logic for DetalleDocs.xaml
    /// </summary>
    public partial class DetalleDocumentos : Window
    {

      // List<PART_ABIERTAS> partidaselecc = new List<PART_ABIERTAS>;
       // public DetalleDocumentos(  List<PART_ABIERTAS> partidaselecc)
        public DetalleDocumentos(string NDoc, string NRef, string RUT, string CodCli, string NomCli, string CeBe, string CtrlCred, string Sociedad
       , string FechaDoc, string FechaVenc, string DiasAtr, string Moneda, string ClaseDoc, string ClaseCta, string CME, string ACC, string Estado
       , string CondPag, string MontoPag, string MontoAbon, string Monto)
        {
            InitializeComponent();
            txtNDoc.Text =  NDoc;
            txtNRef.Text =  NRef;
            txtRUT.Text =RUT;
            txtCodCliente.Text =  CodCli;
            txtNomcliente.Text =  NomCli;
            txtEstado.Text = Estado;
            txtCtrlCred.Text =  CtrlCred;
            txtCeBe.Text =  CeBe;
            txtSociedad.Text =  Sociedad;
            txtFechaDoc.Text =  FechaDoc;
            txtFechaVen.Text =  FechaVenc;
            txtDiasAtraso.Text =  DiasAtr;
            txtMoneda.Text =  Moneda;
            txtClaseDoc.Text =  ClaseDoc;
            txtClaseCta.Text =  ClaseCta;
            txtCME.Text =  CME;
            txtACC.Text = ACC;
            txtCondPago.Text = CondPag;
            txtMontoPag.Text =  MontoPag;
            txtMontoAbon.Text = MontoAbon;
            txtMonto.Text =  Monto;
            this.Topmost = true;
            GC.Collect();
        }

        private void Window_Closed(object sender, RoutedEventArgs e)
        {
          
        }

        private void Window_Closed(object sender, EventArgs e)
        {

        }

      
    }
}
