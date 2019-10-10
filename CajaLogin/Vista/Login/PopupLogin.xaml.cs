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
using CajaIndigo.AppPersistencia.Class.Login.Estructura;
using CajaIndigo.AppPersistencia.Class.Login;
using CajaIndigo.AppPersistencia.Class.AperturaCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.AperturaCaja;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.CierreCaja;
using CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura;
using CajaIndigo.AppPersistencia.Class.UsuariosCaja;
using CajaIndigo;



namespace CajaIndigo
{
    /// <summary>
    /// Interaction logic for PopupLogin.xaml
    /// </summary>
    /// 
    public partial class PopupLogin : Window
    {
        UsuariosCaja usuariocaja = new UsuariosCaja();
        LOG_APERTURA log = new LOG_APERTURA();

        List<CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA> user = new List<CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA>();
        
        CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA userobject = new CajaIndigo.AppPersistencia.Class.UsuariosCaja.Estructura.USR_CAJA();
        
        string Moned = string.Empty;

        Vista.Menu.MenuCaja FrmMenu;

        public PopupLogin(string usuariologg, string pass, string temporal, List<string> listacajas,List<string> sucursales, List<string> pais, List<string> monedas) 
        {
            InitializeComponent();
            textBlock1.Content = usuariologg;
            lblPassWord.Content = pass;
            //Llena combobox, de tener un solo valor el combobox lo selecciona por defecto
            comboBox1.ItemsSource = listacajas;
            comboBox2.ItemsSource = pais;
            comboBox3.ItemsSource = monedas;

            if (comboBox1.Items.Count == 1)
            {
                comboBox1.SelectedIndex = 0;
            }
            if (comboBox2.Items.Count == 1)
            {
                comboBox2.SelectedIndex = 0;
            }
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
                var comboBox = sender as ComboBox;
                string value = comboBox1.SelectedItem as string;
                string VALUE2 = comboBox2.SelectedItem as string;
                string VALUE3 = comboBox3.SelectedItem as string;
                string idcaja = "";
                string nomcaja = "";
                List<string> moneda = new List<string>();
                for (int i = 0; i < comboBox3.Items.Count; i++)
                {
                    moneda.Add(Convert.ToString(comboBox3.Items.GetItemAt(i)));
                    Moned = moneda.ToString();
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

                            userobject.WAERS = Convert.ToString(comboBox3.Text);

                            user.Add(userobject);

                            usuariocaja.usuarioscaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtTemporal.Text, Convert.ToString(textBlock2.Text), user, Environment.MachineName);
                            string Mensaje = usuariocaja.errormessage;
                            string str = "";
                            str = usuariocaja.status;
                            //***RFC Apertura de Caja
                            switch (str)
                            {
                                case "S": //Apertura de caja exitosa
                                    {
                                        logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Apertura caja: " + usuariocaja.Mensaje);

                                        if (txtTemporal.Text != "X")
                                        {
                                            FrmMenu = new Vista.Menu.MenuCaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), Convert.ToString(textBlock1.Content), idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);
                                        }
                                        else
                                        {
                                            FrmMenu = new Vista.Menu.MenuCaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), Convert.ToString(textBlock1.Content), idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);
                                        }

                                        FrmMenu.txtIdSistema.Text = txtIdSistema.Text;
                                        FrmMenu.txtInstancia.Text = txtInstancia.Text;
                                        FrmMenu.txtMandante.Text = txtMandante.Text;
                                        FrmMenu.txtSapRouter.Text = txtSapRouter.Text;
                                        FrmMenu.txtServer.Text = txtServer.Text;
                                        FrmMenu.txtIdioma.Text = txtIdioma.Text;
                                        FrmMenu.idcaja.Text = idcaja;
                                        FrmMenu.NomCaja.Text = nomcaja;
                                        FrmMenu.SociedCaja.Text = usuariocaja.Sociedad;
                                        FrmMenu.PaisCaja.Text = Convert.ToString(comboBox2.Text);
                                        //FrmMenu.MonedCaja.Text = Convert.ToString(moneda);
                                        FrmMenu.usercajaLog.Text = Convert.ToString(usuariocaja.LogApert);
                                        FrmMenu.MonedaCaja.Text = Convert.ToString(comboBox3.Text);
                                        FrmMenu.PassUserCaja.Text = Convert.ToString(lblPassWord.Content);
                                        FrmMenu.UsuarioCaja.Text = Convert.ToString(textBlock1.Content);


                                        FrmMenu.Owner = this.Owner;
                                        this.Width = 140;
                                        FrmMenu.Show();
                                        this.Visibility = Visibility.Collapsed;

                                        break;
                                    }
                                case "E": //Acceso a caja por usuario distinto al que realizo la apertura
                                    {
                                        logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Acceso temporal a caja: " + usuariocaja.Mensaje);
                                        MessageBox.Show("No puede acceder como usuario principal" + usuariocaja.Mensaje + "-" + usuariocaja.status);
                                        //FORM CIERRA Y ABRE DE NUEVO VENTANA DE LOGIN 
                                        MainWindow window = Window.GetWindow(this.Owner) as MainWindow;
                                        if (window != null)
                                        {
                                            this.Close();
                                            window.Visibility = Visibility.Visible;
                                        }
                                        break;
                                    }
                                case "W": //Caja no cerrada la fecha anterior
                                    {
                                        logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Apertura caja fallida: " + usuariocaja.Mensaje);

                                        //***RFC cierre de Caja
                                        CierreCaja cierrecaja = new CierreCaja();
                                        cierrecaja.cierreTempo(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, idcaja, Convert.ToString(comboBox2.Text), "5000", textBlock2.Text, "Probando 1", "Probando 2");
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
                }
            }
            //}
            //   else
            //   {
            //    MessageBox.Show("Ingrese el país para la apertura de caja");
            //   }
            ////}
            //   else
            //   {
            //    MessageBox.Show("Ingrese la moneda para la apertura de caja");
            //   }
            //}
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), Convert.ToString(comboBox1.SelectedItem), "", ex.Message + ex.StackTrace);
            }          
        }
        private void btnInicCaja_Click(object sender, RoutedEventArgs e)
        {
            try
            {
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
                            userobject.WAERS = Convert.ToString(comboBox3.Text);

                            user.Add(userobject);

                            usuariocaja.usuarioscaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, txtTemporal.Text,Convert.ToString(textBlock2.Text), user,Environment.MachineName);
                            string Mensaje = usuariocaja.errormessage;
                            string str = "";
                            str = usuariocaja.status;
                            

                            ////***RFC Apertura de Caja
                            switch (str)
                            {
                                case "S": //Apertura de caja exitosa
                                    {
                                        logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Apertura caja: " + usuariocaja.Mensaje);
                                  
                                        if (txtTemporal.Text != "X")
                                        {
                                           FrmMenu = new Vista.Menu.MenuCaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), Convert.ToString(textBlock1.Content), idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);                                        
                                        }
                                        else
                                        {
                                            FrmMenu = new Vista.Menu.MenuCaja(Convert.ToString(textBlock1.Content), Convert.ToString(lblPassWord.Content), Convert.ToString(textBlock1.Content), idcaja, nomcaja, usuariocaja.Sociedad, moneda, Convert.ToString(comboBox2.Text), Convert.ToDouble(textBlock2.Text), usuariocaja.LogApert);  
                                        }

                                        FrmMenu.txtIdSistema.Text = txtIdSistema.Text;
                                        FrmMenu.txtInstancia.Text = txtInstancia.Text;
                                        FrmMenu.txtMandante.Text = txtMandante.Text;
                                        FrmMenu.txtSapRouter.Text = txtSapRouter.Text;
                                        FrmMenu.txtServer.Text = txtServer.Text;
                                        FrmMenu.txtIdioma.Text = txtIdioma.Text;
                                        FrmMenu.idcaja.Text = idcaja;
                                        FrmMenu.NomCaja.Text = nomcaja;
                                        FrmMenu.SociedCaja.Text = usuariocaja.Sociedad;
                                        //FrmMenu.MonedCaja.Text = Convert.ToString(moneda[0]);
                                        FrmMenu.usercajaLog.Text = Convert.ToString(usuariocaja.LogApert);
                                        FrmMenu.MonedaCaja.Text = Convert.ToString(comboBox3.Text);
                                        FrmMenu.PassUserCaja.Text = Convert.ToString(lblPassWord.Content);
                                        FrmMenu.UsuarioCaja.Text = Convert.ToString(textBlock1.Content);

                                        FrmMenu.Owner = this.Owner;
                                        this.Width = 140;
                                        FrmMenu.Show();
                                        this.Visibility = Visibility.Collapsed;
                                       
                                        break;
                                    }
                                case "E": //Acceso a caja por usuario distinto al que realizo la apertura
                                    {
                                        logCaja.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), idcaja, nomcaja, "Acceso temporal a caja: " + usuariocaja.Mensaje);
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
                LogCajaIndigo.EscribeLogCaja(System.DateTime.Now, Convert.ToString(textBlock1.Content), Convert.ToString(comboBox1.SelectedItem), "", ex.Message + ex.StackTrace);

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
