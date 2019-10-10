using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Deployment.Internal;
using System.Deployment;
using System.Reflection;
using System.Diagnostics;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Login.Estructura;
using CajaIndu.AppPersistencia.Class.Login;


namespace CajaIndu 
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    ///  
    public partial class MainWindow : Window
    {

        public string UsuarioLoggeado;
        public MainWindow()
        {
            InitializeComponent();
           // lblVersion.Content = "Caja Indumotora V-" + System.Windows.Forms.Application.ProductVersion;
           // Version version = Assembly.GetExecutingAssembly().GetName().Version;
           // label10.Content =  Version: {0}.{1}.{2}.{3};
           // label10.Content = String.Format(Convert.ToString(label10.Content), version.Major, version.Minor, version.Build, version.Revision);
           //// label10.Content = "Caja Indumotora V-" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
           // label10.Visibility = Visibility.Collapsed;

        }

        //public Version AssemblyVersion 
        //{
        //    get
        //    {
        //        return ApplicationDeployment.CurrentDeployment.CurrentVersion;
        //    }
        //}

        private void CajaLogin_Loaded(object sender, RoutedEventArgs e)
        {
            lblVersion.Content = "Caja Indumotora V-" + System.Windows.Forms.Application.ProductVersion;
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            label10.Text = String.Format(Convert.ToString(label10.Text), version.Major, version.Minor, version.Build, version.Revision);
            label10.Visibility = Visibility.Collapsed;

        }
        private void button1_GotFocus(object sender, RoutedEventArgs e)
        {

        }

        private void button1_GotFocus_1(object sender, RoutedEventArgs e)
        {

        }

        private void textUser_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                GDConfigSAP.Visibility = Visibility.Collapsed;
                label2.Content = "";
                string user = textUser.Text.ToUpper();
                string clave = passwordBox.Password;
                string s = "";
                string temporal = "";
                List<string> listadocajas = new List<string>();
                List<string> listadosucursales = new List<string>();
                List<string> listadopaises = new List<string>();
                List<string> listadomonedas = new List<string>();
                if (checkBox1.IsChecked == true)
                {
                    temporal = "X";
                }
                //***RFC Login de usuario
                LoginSAP login = new LoginSAP();
                login.datoslogin(user, clave, txtIdSistema.Text, txtInstancia.Text, txtMandante.Text, txtSapRouter.Text, txtServer.Text, txtIdioma.Text, temporal, Environment.MachineName);

               
                //De ser positivo el loggin a SAP, se procede a llenar los datos asociados al usuario (país, moneda, caja, sucursal)
                if (login.ObjDatosLogin.Count > 0)
                {
                    for (int i = 0; i <= login.ObjDatosLogin.Count - 1; i++)
                    {
                        s = login.ObjDatosLogin[i].ID_CAJA;
                        listadocajas.Add(Convert.ToString(login.ObjDatosLogin[i].ID_CAJA) + "-" + Convert.ToString(login.ObjDatosLogin[i].NOM_CAJA));
                        listadosucursales.Add(Convert.ToString(login.ObjDatosLogin[i].NOM_CAJA));
                        if (!listadopaises.Contains(Convert.ToString(login.ObjDatosLogin[i].LAND)))
                        {
                            listadopaises.Add(Convert.ToString(login.ObjDatosLogin[i].LAND));
                        }
                        if (!listadomonedas.Contains(Convert.ToString(login.ObjDatosLogin[i].MONEDA)))
                        {
                            listadomonedas.Add(Convert.ToString(login.ObjDatosLogin[i].MONEDA));
                        }
                    }
                }

                if (login.ObjDatosLogin.Count != 0)
                {
                    //LLama a formulario de ingresos de datos iniciales de apertura de caja
                    PopupLogin frm = new PopupLogin(user, clave, temporal, listadocajas, listadosucursales, listadopaises, listadomonedas);
                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, user, "", "", "Login a SAP exitoso: " + login.errormessage);
                    textUser.Text = "";
                    passwordBox.Password = "";
                    frm.txtIdSistema.Text = txtIdSistema.Text;
                    frm.txtInstancia.Text = txtInstancia.Text;
                    frm.txtMandante.Text = txtMandante.Text;
                    frm.txtSapRouter.Text = txtSapRouter.Text;
                    frm.txtServer.Text = txtServer.Text;
                    frm.txtIdioma.Text = txtIdioma.Text;
                    frm.Owner = this; 
                    frm.Show();
                    //this.Close();
                    this.Visibility = Visibility.Collapsed;
                }
                else
                {
                    //Si falla la conexión a SAP limpia el user y password y registra el evento en el log
                    label2.Content = login.errormessage;
                    LogCajaIndu logtxt = new LogCajaIndu();
                    LogCajaIndu.EscribeLogCajaIndumotora(System.DateTime.Now, user, "", "", "Error : Login a SAP. " + login.errormessage);
                    textUser.Text = "";
                    passwordBox.Password = "";
                }
                GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + ex.StackTrace);
                System.Windows.MessageBox.Show(ex.Message );
            }
        }

        private void CajaLogin_Closed(object sender, EventArgs e)
        {     
           foreach (System.Diagnostics.Process myProc in System.Diagnostics.Process.GetProcesses())
           {
               if (myProc.ProcessName == "CajaIndu")
               {
                   myProc.Kill();
               }
           } 
        }

        private void btnConfig_Click(object sender, RoutedEventArgs e)
        {
            GDConfigSAP.Visibility = Visibility.Visible;
        }

        private void CajaLogin_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ;
        }



     
       


    }
}
