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
using System.IO;


namespace CajaIndu
{
    /// <summary>
    /// Interaction logic for PDFViewer.xaml
    /// </summary>
    public partial class PDFViewer : Window
    {
        public PDFViewer()
        {
            InitializeComponent();
            //GBReimpresionFiscal.Visibility = Visibility.Collapsed;
            //webBrowser1.Navigate("about:blank");
            //this.Owner = System.Windows.Window(CajaIndu.Comprobante);//.UidProperty(CajaIndu.Comprobante);
            string Owner = Convert.ToString(this.Owner);
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            webBrowser1.Navigate("about:blank");
            webBrowser1.Dispose();

            //File.Delete(txtArchivo.Text);
            //File.Delete(txtArchivoNuevo.Text);
            try
            {
                Comprobante window = Window.GetWindow(this.Owner) as Comprobante;
                if (window != null)
                {
                    //this.Close();
                    window.Close();
                }
                GC.Collect();
            }
            catch(Exception  ex)
            {
               Console.WriteLine(ex.Message + ex.StackTrace);                       
            }
            GC.Collect();
        }
    }
}
