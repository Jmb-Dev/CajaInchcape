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
using CajaIndu;


namespace CajaIndu
{
    /// <summary>
    /// Interaction logic for ResumenViasPago.xaml
    /// </summary>
    
    public partial class ResumenViasPago : Window
    {
       
        public ResumenViasPago()
        {
            InitializeComponent();
          //  DGResumenViasPago.ItemsSource = ListViasPag;
        }

        //FUNCION QUE ACTUALIZA LAS VIAS DE PAGO LEYENDO LA INFORMACION DE LA GRILLA Y ACTUALIZANDO LA GRILLA DGCHEQUES 
        //(QUE GUARDA LA INFO DE LAS VIAS DE PAGO AL REALIZAR EL PAGO POR LA RFC)
        public void ActualizaViasPago()
        {
            List<DetalleViasPago> ListViasPag = new List<DetalleViasPago>(); 
            for (int i = 0; i < DGResumenViasPago.Items.Count; i++)
            {

                ListViasPag.Add(DGResumenViasPago.Items.CurrentItem as DetalleViasPago);
                DGResumenViasPago.Items.MoveCurrentToNext();

            }
            //FORM RESUMENVIASPAGO ENVIA LA INFORMACION AL DGCHEQUES 
            PagosDocumentos window = Window.GetWindow(this.Owner) as PagosDocumentos;
            if (window != null)
            {
                this.Close();
                window.DGCheque.ItemsSource = ListViasPag;
             }
            Close();
        }
      
        //EVENTO DE CERRAR EL FORM
        private void Window_Closed(object sender, EventArgs e)
        {
            ActualizaViasPago();
        }
    }
}
