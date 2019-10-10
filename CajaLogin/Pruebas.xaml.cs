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
using System.Windows.Navigation;
using System.Windows.Shapes;
using CajaIndigo.AppPersistencia.Class.CierreCaja;
using CajaIndigo.AppPersistencia.Class.CierreCaja.Estructura;
using Org.BouncyCastle.Asn1.Crmf;

namespace CajaIndigo
{
    /// <summary>
    /// Interaction logic for Pruebas.xaml
    /// </summary>
    public partial class Pruebas : Page
    {


        public Pruebas()
        {
            InitializeComponent();
            cargarGrid();
        }

        private void cargarGrid()
        {
            string UserCaja = "FPONCE";
            string PassCaja = "FELIPE06";
            string IdCaja = "";
            string NombCaja = "";
            string SociedadCaja = "EI15";
            string MonedCaja = "CLP";
            string PaisCja = "CL";
            string Monto = "0";
            string IdSistema = "INQ";
            string Instancia = "00";
            string mandante = "200";
            string SapRouter = "";
            string server = "10.9.100.168";
            string idioma = "ES";
            double monto;
            double monto2;
            string Valor2 = string.Empty;
            string moneda = string.Empty;

            CierreCaja Cierre = new CierreCaja();
           

            Cierre.OtrasMonedas(UserCaja, PassCaja, IdSistema, Instancia, mandante, SapRouter, server, idioma, IdCaja, PaisCja);
            
            if (Cierre.MonedExtr.Count > 0)
            {
                
                    prueba2.ShowGridLines = true;
                    //prueba2.Background = new SolidColorBrush(Colors.Gray);
                    // Create Columns
                    ColumnDefinition gridCol1 = new ColumnDefinition();
                    ColumnDefinition gridCol2 = new ColumnDefinition();
                    prueba2.ColumnDefinitions.Add(gridCol1);
                    prueba2.ColumnDefinitions.Add(gridCol2);

                    // Create Rows
                    RowDefinition gridRow1 = new RowDefinition();
                    gridRow1.Height = new GridLength(45);
                    RowDefinition gridRow2 = new RowDefinition();
                    gridRow2.Height = new GridLength(45);
                    prueba2.RowDefinitions.Add(gridRow1);
                    prueba2.RowDefinitions.Add(gridRow2);

                    for (int i = 0; i < Cierre.MonedExtr.Count(); i++)
                    {
                        // Crear Caja de Textos
                        TextBox MyTextBox = new System.Windows.Controls.TextBox();
                        TextBox MyTextBox2 = new System.Windows.Controls.TextBox();
                        MyTextBox.Text = Cierre.MonedExtr[i].MONEDA;
                        MyTextBox.Name = Cierre.MonedExtr[i].MONEDA;
                        MyTextBox.FontSize = 12;
                        MyTextBox.VerticalAlignment = VerticalAlignment.Top;
                        MyTextBox.IsEnabled = false;
                        MyTextBox.FontWeight = FontWeights.Bold;
                        Grid.SetRow(MyTextBox, i);
                        Grid.SetColumn(MyTextBox, 0);
                            
                        
                        prueba2.Children.Add(MyTextBox);
                        Grid.SetRow(MyTextBox2, i);
                        Grid.SetColumn(MyTextBox2, 1);
                        prueba2.Children.Add(MyTextBox2);

                        //TextBox MyTextBox2 = new System.Windows.Controls.TextBox();
                        //MyTextBox2.Text = "";
                        //MyTextBox2.FontSize = 14;
                        //MyTextBox2.FontWeight = FontWeights.Bold;
                        //MyTextBox2.Foreground = new SolidColorBrush(Colors.Green);
                        //MyTextBox2.VerticalAlignment = VerticalAlignment.Top;
                        //Grid.SetColumn(MyTextBox2, i);
                        //Grid.SetColumn(MyTextBox2, i);
                        //prueba2.Children.Add(MyTextBox2);
                    }
                        
                }
            }
        }
    }
