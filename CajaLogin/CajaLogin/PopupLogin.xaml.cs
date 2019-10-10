using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using CajaIndu.AppPersistencia.Class.Login.Estructura;
using CajaIndu.AppPersistencia.Class.Login;
using CajaIndu.AppPersistencia.Class.AperturaCaja.Estructura;
using CajaIndu.AppPersistencia.Class.AperturaCaja;
using CajaIndu.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndu.AppPersistencia.Class.CierreCaja;
using CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura;
using CajaIndu.AppPersistencia.Class.UsuariosCaja;



namespace CajaIndu
{
    /// <summary>
    /// Interaction logic for PopupLogin.xaml
    /// </summary>
    /// 
    public partial class PopupLogin : Window
    {
        UsuariosCaja usuariocaja = new UsuariosCaja();
        List<CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA> user = new List<CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA>();
        CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA userobject = new CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA();
 
        public PopupLogin(string usuariologg, string pass, string temporal, List<string> listacajas,List<string> sucursales, List<string> pais, List<string> monedas) 
        {
            InitializeComponent();
            textBlock1.Content = usuariologg;
            lblPassWord.Content = pass;
            //Llena combobox, de tener un solo valor el combobox lo selecciona por defecto
            comboBox1.ItemsSource = listacajas;
            if (comboBox1.Items.Count == 1)
            {
                comboBox1.SelectedIndex = 0;
            }
            comboBox2.ItemsSource = pais;
            if (comboBox2.Items.Count == 1)
            {
                comboBox2.SelectedIndex = 0;
            }
            comboBox3.ItemsSource = monedas;
            if (comboBox3.Items.Count == 1)
            {
                comboBox3.SelectedIndex = 0;
            }
            txtTemporal.Text = temporal;



           

           
        }

        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                // ... Get the ComboBox.
                var comboBox = sender as ComboBox;

                // ... Set SelectedItem as como Caja recaudadora
                string value = comboBox1.SelectedItem as string;
                string idcaja = "";
                string nomcaja = "";
                List<string> moneda = new List<string>();
                for (int i = 0; i < comboBox3.Items.Count; i++)
                {
                    moneda.Add(Convert.ToString(comboBox3.Items.GetItemAt(i)));
                }
                int size = value.IndexOf("-");
                idcaja = value.Substring(0, size);
                size += 1;
                nomcaja = value.Substring(size, value.Length - size);
                //Validación de moneda
                if ( Convert.ToString(comboBox3.Text) != "")
                {
                    //Validacion de pais
                 if ( Convert.ToString(comboBox2.Text) != "")
                    {
                     //Validacion de monto inicial de apertura de caja
                    if  (Convert.ToString(textBlock2.Text) != "")
                    {
                        userobject.ID_CAJA = idcaja;
                        userobject.LAND = Convert.ToString(comboBox2.Text);
                        userobject.NOM_CAJA = nomcaja;
                        if ( txtTemporal.Text == "X")
                        {
                            userobject.TIPO_USUARIO = "T";
                        }
                        else
                        {
                            userobject.TIPO_USUARIO = "P";
                        }
                        userobject.USUARIO = Convert.ToString(textBlock1.Content);

                        userobject.WAERS = Convert.ToString(comboBox3.Text);



                        user.Add(userobject);

                        usuariocaja.usuarioscaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtTemporal.Text, Convert.ToString(textBlock2.Text), user, Environment.MachineName);
                        string Mensaje = usuariocaja.errormessage;
                        string str = "";
                        str = usuariocaja.status;

                        
                        //***RFC Apertura de Caja
                        //AperturaCaja aperturacaja = new AperturaCaja();
                        //aperturacaja.aperturacaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), idcaja, Convert.ToString(textBlock2.Text),Convert.ToString(comboBox2.Text),Convert.ToString(comboBox3.Text), "A");
                        //string str = "";
                        //str = aperturacaja.status;
                        switch (str)
                        {
                            case "S": //Apertura de caja exitosa
                                {
                                   
                                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Apertura caja: " + usuariocaja.Mensaje);
                                    
                                    PagosDocumentos frm;
                                    if (txtTemporal.Text != "X")
                                    {
                                        frm = new PagosDocumentos(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), Convert.ToString(textBlock1.Content), idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);
                                    }
                                    else
                                    {
                                        frm = new PagosDocumentos(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), usuariocaja.cajeroresp, idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);
                                    }
                                    frm.txtIdSistema.Text = txtIdSistema.Text;
                                    frm.txtInstancia.Text = txtInstancia.Text;
                                    frm.txtMandante.Text = txtMandante.Text;
                                    frm.txtSapRouter.Text = txtSapRouter.Text;
                                    frm.txtServer.Text = txtServer.Text;
                                    frm.txtIdioma.Text = txtIdioma.Text;
                                    frm.Owner = this.Owner;
                                    this.Width = 140;
                                    frm.Show();
                                    this.Visibility = Visibility.Collapsed;
                                    
                                    break;
                                }
                            case "E": //Acceso a caja por usuario distinto al que realizo la apertura
                                {
                                   LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Acceso temporal a caja: " + usuariocaja.Mensaje);
                                   MessageBox.Show("No puede acceder como usuario principal" + usuariocaja.Mensaje + "-" + usuariocaja.status);
                                   //FORM CIERRA Y ABRE DE NUEVO VENTANA DE LOGIN 
                                   MainWindow window = Window.GetWindow(this.Owner) as MainWindow;
                                   if (window != null)
                                   {
                                       this.Close();
                                       //window.Show();
                                       window.Visibility = Visibility.Visible;
                                   }
                                   
                                    break;
                                   
                                }
                            case "W": //Caja no cerrada la fecha anterior
                                {
                                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Apertura caja fallida: " + usuariocaja.Mensaje);
                       
                                   // MessageBox.Show(aperturacaja.message + " " + aperturacaja.errormessage);                                 
                                    //***RFC cierre de Caja
                                    CierreCaja cierrecaja = new CierreCaja();
                                    cierrecaja.cierrecaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, idcaja, Convert.ToString(comboBox2.Text), "5000", textBlock2.Text, "Probando 1", "Probando 2");
                                    MessageBox.Show(cierrecaja.T_Retorno[0].MESSAGE.ToString());
                                    break;
                                }
                            default:
                                {
                                    MessageBox.Show(usuariocaja.Mensaje + " " + Mensaje);
                                    // MessageBox.Show(aperturacaja.errormessage);
                                    break;
                                }
                        }
                    }
                    else
                    {
                     MessageBox.Show("Ingrese un monto para la apertura de caja");
                    }
                 }
                else
                {
                 MessageBox.Show("Ingrese el país para la apertura de caja");
                }
             }
                else
                {
                 MessageBox.Show("Ingrese la moneda para la apertura de caja");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), Convert.ToString(comboBox1.SelectedItem), "",  ex.Message + ex.StackTrace);
                       
            }
            
        }

        private void btnInicCaja_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // ... Get the ComboBox.
              //  var comboBox = sender as ComboBox;

                // ... Set SelectedItem as como Caja recaudadora

                if (comboBox1.SelectedItem != null)
                {
                string value = comboBox1.SelectedItem as string;
                string idcaja = "";
                string nomcaja = "";
                List<string> moneda = new List<string>();
                for (int i = 0; i < comboBox3.Items.Count; i++)
                {
                    moneda.Add(Convert.ToString(comboBox3.Items.GetItemAt(i)));
                }


                int size = value.IndexOf("-");
                idcaja = value.Substring(0, size);
                size += 1;
                nomcaja = value.Substring(size, value.Length - size);
                //Validación de moneda
                if (Convert.ToString(comboBox3.Text) != "")
                {
                    //Validacion de pais
                    if (Convert.ToString(comboBox2.Text) != "")
                    {
                        //Validacion de monto inicial de apertura de caja
                        if (Convert.ToString(textBlock2.Text) != "")
                        {
                            userobject.ID_CAJA = idcaja;
                            userobject.LAND = Convert.ToString(comboBox2.Text);
                            userobject.NOM_CAJA = nomcaja;
                            if (txtTemporal.Text == "X")
                            {
                                userobject.TIPO_USUARIO = "T";
                            }
                            else
                            {
                                userobject.TIPO_USUARIO = "P";
                            }
                            userobject.USUARIO = Convert.ToString(textBlock1.Content);

                            //userobject.WAERS = moneda;



                            user.Add(userobject);

                            usuariocaja.usuarioscaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtTemporal.Text,Convert.ToString(textBlock2.Text), user,Environment.MachineName);
                            string Mensaje = usuariocaja.errormessage;
                            string str = "";
                            str = usuariocaja.status;

                            ////***RFC Apertura de Caja
                            //AperturaCaja aperturacaja = new AperturaCaja();
                            //aperturacaja.aperturacaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), idcaja, Convert.ToString(textBlock2.Text), Convert.ToString(comboBox2.Text), Convert.ToString(comboBox3.Text), "A");
                            //string str = "";
                            //str = aperturacaja.status;
                            switch (str)
                            {
                                case "S": //Apertura de caja exitosa
                                    {

                                        LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Apertura caja: " + usuariocaja.Mensaje);

                                        PagosDocumentos frm;
                                        if (txtTemporal.Text != "X")
                                        {
                                            frm = new PagosDocumentos(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), Convert.ToString(textBlock1.Content), idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);
                                        }
                                        else
                                        {
                                           
                                            frm = new PagosDocumentos(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), usuariocaja.cajeroresp, idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);
                                        }
                                        frm.txtIdSistema.Text = txtIdSistema.Text;
                                        frm.txtInstancia.Text = txtInstancia.Text;
                                        frm.txtMandante.Text = txtMandante.Text;
                                        frm.txtSapRouter.Text = txtSapRouter.Text;
                                        frm.txtServer.Text = txtServer.Text;
                                        frm.Owner = this.Owner;
                                        this.Width = 140;
                                        frm.Show();
                                        
                                       
                                        this.Visibility = Visibility.Collapsed;
                                        

                                        break;
                                    }
                                case "E": //Acceso a caja por usuario distinto al que realizo la apertura
                                    {
                                        LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Acceso temporal a caja: " + usuariocaja.Mensaje);
                                        MessageBox.Show("No puede acceder como usuario principal" + usuariocaja.Mensaje + "-" + usuariocaja.status);
                                        break;
                                        MainWindow window = Window.GetWindow(this.Owner) as MainWindow;
                                        if (window != null)
                                        {
                                            this.Close();
                                            window.Visibility = Visibility.Visible;
                                        }
                                    }
                                default:
                                    {
                                        MessageBox.Show(usuariocaja.Mensaje + " " + usuariocaja.errormessage);
                                        // MessageBox.Show(aperturacaja.errormessage);
                                        break;
                                    }
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ingrese un monto para la apertura de caja");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Ingrese el país para la apertura de caja");
                    }
                }
                else
                {
                    MessageBox.Show("Ingrese la moneda para la apertura de caja");
                }
            }
            else
            {
                MessageBox.Show("Seleccione la caja donde desea iniciar la sesión de trabajo");
            }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, Convert.ToString(textBlock1.Content), Convert.ToString(comboBox1.SelectedItem), "", ex.Message + ex.StackTrace);

            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow window = System.Windows.Window.GetWindow(this.Owner) as MainWindow;
            
            if (window != null)
            {
                if (this.Width != 140)
                {
                    window.Visibility = Visibility.Visible;
                }
                this.Close();
      
               
            }

        }





   
    }
}
