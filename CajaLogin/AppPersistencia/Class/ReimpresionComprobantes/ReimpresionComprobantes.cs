using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes.Estructura;
using CajaIndigo.AppPersistencia.Class.BusquedaReimpresiones.Estructura;

namespace CajaIndigo.AppPersistencia.Class.ReimpresionComprobantes
{
    class ReimpresionComprobantes
    {
        public string NumDocCont = "";
        public string pagomessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public int id_error = 0;
        ConexSAP connectorSap = new ConexSAP();
        //public List<ESTATUS> T_Retorno = new List<ESTATUS>();
        public List<DOCUMENTOS> DatosCabecera = new List<DOCUMENTOS>();
        public List<VIAS_PAGO2> DatosDetalle = new List<VIAS_PAGO2>();
        public List<DATOS_DOCUMENTOS> DatosDocumentos = new List<DATOS_DOCUMENTOS>();
        public List<DATOS_VP> DatosViaPago = new List<DATOS_VP>();
        public List<DATOS_CAJA> DatosCaja = new List<DATOS_CAJA>();
        public List<DATOS_CLIENTES> DatosCliente = new List<DATOS_CLIENTES>();
        public List<INFO_SOC> DatosEmpresa = new List<INFO_SOC>();
        FormatoMonedas Formato = new FormatoMonedas();


        public void reimprcomprobantes(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, List<VIAS_PAGO2> P_VIASPAGO, List<DOCUMENTOS> P_DOCSAPAGAR)
        {
            try
            {
                DatosCabecera.Clear();
                DatosDetalle.Clear();
                DatosDocumentos.Clear();
                DatosViaPago.Clear();
                DatosCaja.Clear();
                DatosCliente.Clear();
                DatosEmpresa.Clear();
                IRfcStructure lt_DATOS_CAJA;
                IRfcStructure lt_DATOS_CLIENTES;
                IRfcTable lt_DATOS_DOCUMENTOS;
                IRfcTable lt_DATOS_VP;
                IRfcTable lt_DATOSEMPRESA;
                //DATOS_CAJA DATOS_CAJA_resp;
                DATOS_CLIENTES datosclientes;
                DATOS_CAJA datoscaja;
                DATOS_DOCUMENTOS datosdocumentos;
                DATOS_VP datosvp;
                INFO_SOC datosempresa;
                //Conexion a SAP
                connectorSap.idioma = P_IDIOMA;
                connectorSap.idSistema = P_IDSISTEMA;
                connectorSap.instancia = P_INSTANCIA;
                connectorSap.mandante = P_MANDANTE;
                connectorSap.paswr = P_PASSWORD;
                connectorSap.sapRouter = P_SAPROUTER;
                connectorSap.user = P_UNAME;
                connectorSap.server = P_SERVER;

                string retval = connectorSap.connectionsSAP();
                
                //Si el valor de retorno es nulo o vacio, hay conexion a SAP y la RFC trae datos   
                if (string.IsNullOrEmpty(retval))
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                   
                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_FM_REC_IMP_COMPROBANTE");
                    try
                    {

                        IRfcTable GralDat = BapiGetUser.GetTable("VIAS_PAGO");
                        for (var i = 0; i < P_VIASPAGO.Count; i++)
                        {
                            GralDat.Append();
                            GralDat.SetValue("MANDT", P_VIASPAGO[i].MANDT);
                            GralDat.SetValue("LAND", P_VIASPAGO[i].LAND);
                            GralDat.SetValue("ID_COMPROBANTE", P_VIASPAGO[i].ID_COMPROBANTE);
                            GralDat.SetValue("ID_DETALLE", P_VIASPAGO[i].ID_DETALLE);
                            GralDat.SetValue("VIA_PAGO", P_VIASPAGO[i].VIA_PAGO);
                            GralDat.SetValue("MONTO", P_VIASPAGO[i].MONTO ); 
                            GralDat.SetValue("MONEDA", P_VIASPAGO[i].MONEDA);
                            if (P_VIASPAGO[i].BANCO != "")
                            {
                                GralDat.SetValue("BANCO", P_VIASPAGO[i].BANCO.Substring(0, 3));
                            }
                            else
                            {
                                GralDat.SetValue("BANCO", P_VIASPAGO[i].BANCO);
                            }
                            GralDat.SetValue("EMISOR", P_VIASPAGO[i].EMISOR);
                            GralDat.SetValue("NUM_CHEQUE", P_VIASPAGO[i].NUM_CHEQUE);
                            GralDat.SetValue("COD_AUTORIZACION", P_VIASPAGO[i].COD_AUTORIZACION);
                            GralDat.SetValue("NUM_CUOTAS", P_VIASPAGO[i].NUM_CUOTAS);
                            if (P_VIASPAGO[i].FECHA_VENC != "0000-00-00")
                            {
                                GralDat.SetValue("FECHA_VENC", Convert.ToDateTime(P_VIASPAGO[i].FECHA_VENC));
                            }
                            GralDat.SetValue("TEXTO_POSICION", P_VIASPAGO[i].TEXTO_POSICION);
                            GralDat.SetValue("ANEXO", P_VIASPAGO[i].ANEXO);
                            GralDat.SetValue("SUCURSAL", P_VIASPAGO[i].SUCURSAL);
                            GralDat.SetValue("NUM_CUENTA", P_VIASPAGO[i].NUM_CUENTA);
                            GralDat.SetValue("NUM_TARJETA", P_VIASPAGO[i].NUM_TARJETA);
                            GralDat.SetValue("NUM_VALE_VISTA", P_VIASPAGO[i].NUM_VALE_VISTA);
                            GralDat.SetValue("PATENTE", P_VIASPAGO[i].PATENTE);
                            GralDat.SetValue("NUM_VENTA", P_VIASPAGO[i].NUM_VENTA);
                            GralDat.SetValue("PAGARE", P_VIASPAGO[i].PAGARE);
                            if (P_VIASPAGO[i].FECHA_EMISION != "0000-00-00")
                            {
                                GralDat.SetValue("FECHA_EMISION", Convert.ToDateTime(P_VIASPAGO[i].FECHA_EMISION));
                            }
                            GralDat.SetValue("NOMBRE_GIRADOR", P_VIASPAGO[i].NOMBRE_GIRADOR);
                            GralDat.SetValue("CARTA_CURSE", P_VIASPAGO[i].CARTA_CURSE);
                            GralDat.SetValue("NUM_TRANSFER", P_VIASPAGO[i].NUM_TRANSFER);
                            GralDat.SetValue("NUM_DEPOSITO", P_VIASPAGO[i].NUM_DEPOSITO);
                            GralDat.SetValue("CTA_BANCO", P_VIASPAGO[i].CTA_BANCO);
                            GralDat.SetValue("IFINAN", P_VIASPAGO[i].IFINAN);
                            GralDat.SetValue("CORRE", P_VIASPAGO[i].CORRE);
                            GralDat.SetValue("ZUONR", P_VIASPAGO[i].ZUONR);
                            GralDat.SetValue("HKONT", P_VIASPAGO[i].HKONT);
                            GralDat.SetValue("PRCTR", P_VIASPAGO[i].PRCTR);
                        }
                        BapiGetUser.SetValue("VIAS_PAGO", GralDat);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0} Exception caught.", ex);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                    try
                    {

                    IRfcTable GralDat2 = BapiGetUser.GetTable("DOCUMENTOS");
                    for (var i = 0; i < P_DOCSAPAGAR.Count; i++)
                    {
                        GralDat2.Append();
                        GralDat2.SetValue("MANDT", P_DOCSAPAGAR[i].MANDT);
                        GralDat2.SetValue("LAND", P_DOCSAPAGAR[i].LAND);
                        GralDat2.SetValue("ID_COMPROBANTE",  P_DOCSAPAGAR[i].ID_COMPROBANTE);
                        GralDat2.SetValue("POSICION", P_DOCSAPAGAR[i].POSICION);
                        GralDat2.SetValue("CLIENTE", P_DOCSAPAGAR[i].CLIENTE);
                        GralDat2.SetValue("TIPO_DOCUMENTO", P_DOCSAPAGAR[i].CLASE_DOC);
                        GralDat2.SetValue("SOCIEDAD", P_DOCSAPAGAR[i].SOCIEDAD);
                        if (i == 0)
                        {
                            NumDocCont = P_DOCSAPAGAR[i].NRO_DOCUMENTO;
                        }
                        GralDat2.SetValue("NRO_DOCUMENTO", P_DOCSAPAGAR[i].NRO_DOCUMENTO);
                        GralDat2.SetValue("NRO_REFERENCIA", P_DOCSAPAGAR[i].NRO_REFERENCIA);
                        GralDat2.SetValue("CAJERO_RESP", P_DOCSAPAGAR[i].CAJERO_RESP);
                        GralDat2.SetValue("CAJERO_GEN", P_DOCSAPAGAR[i].CAJERO_GEN);
                        GralDat2.SetValue("ID_CAJA",P_DOCSAPAGAR[i].ID_CAJA);
                        GralDat2.SetValue("FECHA_COMP", P_DOCSAPAGAR[i].FECHA_COMP);
                        GralDat2.SetValue("HORA", P_DOCSAPAGAR[i].HORA);
                        GralDat2.SetValue("NRO_COMPENSACION", P_DOCSAPAGAR[i].NRO_COMPENSACION);
                        GralDat2.SetValue("TEXTO_CABECERA", P_DOCSAPAGAR[i].TEXTO_CABECERA);
                        GralDat2.SetValue("NULO", P_DOCSAPAGAR[i].NULO);
                        GralDat2.SetValue("USR_ANULADOR", P_DOCSAPAGAR[i].USR_ANULADOR);
                        GralDat2.SetValue("NRO_ANULACION", P_DOCSAPAGAR[i].NRO_ANULACION);
                        GralDat2.SetValue("APROBADOR_ANULA", P_DOCSAPAGAR[i].APROBADOR_ANULA);
                        GralDat2.SetValue("TXT_ANULACION", P_DOCSAPAGAR[i].TXT_ANULACION);
                        GralDat2.SetValue("EXCEPCION",P_DOCSAPAGAR[i].EXCEPCION);
                        if (P_DOCSAPAGAR[i].FECHA_COMP != "0000-00-00")
                        {
                            GralDat2.SetValue("FECHA_DOC", Convert.ToDateTime(P_DOCSAPAGAR[i].FECHA_COMP));
                        }
                        if (P_DOCSAPAGAR[i].FECHA_VENC_DOC != "0000-00-00")
                        {
                            GralDat2.SetValue("FECHA_VENC_DOC", Convert.ToDateTime(P_DOCSAPAGAR[i].FECHA_VENC_DOC));
                        }
                        GralDat2.SetValue("NUM_CUOTA", P_DOCSAPAGAR[i].NUM_CUOTA);
                        GralDat2.SetValue("MONTO_DOC", P_DOCSAPAGAR[i].MONTO_DOC);
                        GralDat2.SetValue("MONTO_DIFERENCIA", P_DOCSAPAGAR[i].MONTO_DIFERENCIA);
                        GralDat2.SetValue("TEXTO_EXCEPCION", P_DOCSAPAGAR[i].TEXTO_EXCEPCION);
                        GralDat2.SetValue("PARCIAL", P_DOCSAPAGAR[i].PARCIAL);
                        GralDat2.SetValue("TIME", P_DOCSAPAGAR[i].TIME);
                        GralDat2.SetValue("APROBADOR_EX", P_DOCSAPAGAR[i].APROBADOR_EX);
                        GralDat2.SetValue("MONEDA", P_DOCSAPAGAR[i].MONEDA);
                        GralDat2.SetValue("CLASE_CUENTA", P_DOCSAPAGAR[i].CLASE_CUENTA);
                        GralDat2.SetValue("CLASE_DOC", P_DOCSAPAGAR[i].CLASE_DOC);
                        GralDat2.SetValue("NUM_CANCELACION", P_DOCSAPAGAR[i].NUM_CANCELACION);
                        GralDat2.SetValue("CME", P_DOCSAPAGAR[i].CME);
                        GralDat2.SetValue("NOTA_VENTA", P_DOCSAPAGAR[i].NOTA_VENTA);
                        GralDat2.SetValue("CEBE", P_DOCSAPAGAR[i].CEBE);
                        GralDat2.SetValue("ACC", P_DOCSAPAGAR[i].ACC);
                    }
                    BapiGetUser.SetValue("DOCUMENTOS", GralDat2);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0} Exception caught.", ex);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                    BapiGetUser.Invoke(SapRfcDestination);
                    //LLenamos los datos que retorna la estructura de la RFC
                    try
                    {
                        lt_DATOS_CAJA = BapiGetUser.GetStructure("DATOS_CAJA");
                        for (int i = 0; i < lt_DATOS_CAJA.Count(); i++)
                        {
                            //lt_DATOS_CAJA.CurrentIndex = i;
                            datoscaja = new DATOS_CAJA();
                            datoscaja.NAME_CAJERO = lt_DATOS_CAJA.GetString("NAME_CAJERO");
                            datoscaja.USUARIO = lt_DATOS_CAJA.GetString("USUARIO");
                            datoscaja.ID_COMPROBANTE = lt_DATOS_CAJA.GetString("ID_COMPROBANTE");
                            datoscaja.NRO_DOCUMENTO = lt_DATOS_CAJA.GetString("NRO_DOCUMENTO");
                            datoscaja.NOM_SOCIEDAD = lt_DATOS_CAJA.GetString("NOM_SOCIEDAD");
                            datoscaja.RUT_SOCIEDAD = lt_DATOS_CAJA.GetString("RUT_SOCIEDAD");
                            datoscaja.NOM_CAJA = lt_DATOS_CAJA.GetString("NOM_CAJA");
                            DatosCaja.Add(datoscaja);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0} Exception caught.", ex);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }
                    try
                    {
                        lt_DATOS_CLIENTES = BapiGetUser.GetStructure("DATOS_CLIENTES");
                        for (int i = 0; i < lt_DATOS_CLIENTES.Count(); i++)
                        {
                            //lt_DATOS_CLIENTES.CurrentIndex = i;
                            datosclientes = new DATOS_CLIENTES();
                            datosclientes.RUT = lt_DATOS_CLIENTES.GetString("RUT");
                            datosclientes.NOMBRE = lt_DATOS_CLIENTES.GetString("NOMBRE");
                            DatosCliente.Add(datosclientes);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0} Exception caught.", ex);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                try
                {
                    lt_DATOS_DOCUMENTOS = BapiGetUser.GetTable("DATOS_DOCUMENTOS");
                    for (int i = 0; i < lt_DATOS_DOCUMENTOS.Count(); i++)
                    {
                        lt_DATOS_DOCUMENTOS.CurrentIndex = i;
                        datosdocumentos = new DATOS_DOCUMENTOS();
                        datosdocumentos.TXT_DOCU = lt_DATOS_DOCUMENTOS.GetString("TXT_DOCU");
                        datosdocumentos.NRO_DOCUMENTO = lt_DATOS_DOCUMENTOS.GetString("NRO_DOCUMENTO");
                        datosdocumentos.FECHA_DOC = lt_DATOS_DOCUMENTOS.GetString("FECHA_DOC");
                        datosdocumentos.FECHA_VENC_DOC = lt_DATOS_DOCUMENTOS.GetString("FECHA_VENC_DOC");
                        datosdocumentos.MONEDA = lt_DATOS_DOCUMENTOS.GetString("MONEDA");
                        string str = datosdocumentos.MONEDA;
                        switch (str)
                        {

                            case "USD":
                                //string Valor3 = lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_MO").Trim();
                                //Valor3 = Valor3.Replace(".", "");
                                //Valor3 = Valor3.Replace(",", "");
                                //decimal ValorAux3 = Convert.ToDecimal(Valor3.Substring(0, Valor3.Length - 2));
                                //string Cualquiernombre3 = string.Format("{0:0,0}", ValorAux3);
                                datosdocumentos.MONTO_DOC_MO = Formato.FormatoMonedaExtranjera(lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_MO").Trim());

                                //string Valor4 = lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_ML").Trim();
                                //Valor4 = Valor4.Replace(".", "");
                                //Valor4 = Valor4.Replace(",", "");
                                //decimal ValorAux4 = Convert.ToDecimal(Valor4.Substring(0, Valor4.Length - 2));
                                //string Cualquiernombre4 = string.Format("{0:0,0}", ValorAux4);
                                datosdocumentos.MONTO_DOC_ML = Formato.FormatoMonedaExtranjera(lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_ML").Trim());
                                break;

                            case "EUR":
                                //string Valor6 = lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_MO").Trim();
                                //Valor6 = Valor6.Replace(".", "");
                                //Valor6 = Valor6.Replace(",", "");
                                //decimal ValorAux6 = Convert.ToDecimal(Valor6.Substring(0, Valor6.Length - 2));
                                //string Cualquiernombre6 = string.Format("{0:0,0}", ValorAux6);
                                datosdocumentos.MONTO_DOC_MO = Formato.FormatoMonedaExtranjera(lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_MO").Trim());

                                //string Valor7 = lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_ML").Trim();
                                //Valor7 = Valor7.Replace(".", "");
                                //Valor7 = Valor7.Replace(",", "");
                                //decimal ValorAux7 = Convert.ToDecimal(Valor7.Substring(0, Valor7.Length - 2));
                                //string Cualquiernombre7 = string.Format("{0:0,0}", ValorAux7);
                                datosdocumentos.MONTO_DOC_ML = Formato.FormatoMonedaExtranjera(lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_ML").Trim());
                                break;
                            case "CLP":
                                datosdocumentos.MONTO_DOC_MO = Formato.FormatoMoneda(lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_MO"));
                                datosdocumentos.MONTO_DOC_ML = Formato.FormatoMoneda(lt_DATOS_DOCUMENTOS.GetString("MONTO_DOC_ML"));
                                break;
                        } 
                        datosdocumentos.PEDIDO = lt_DATOS_DOCUMENTOS.GetString("PEDIDO");
                        DatosDocumentos.Add(datosdocumentos);
                    }
                 }
                 catch (Exception ex)
                 {
                            Console.WriteLine("{0} Exception caught.", ex);
                            System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                 }
                try
                {
                    lt_DATOSEMPRESA = BapiGetUser.GetTable("INFO_SOC");
                    for (int i = 0; i < lt_DATOSEMPRESA.Count(); i++)
                    {
                        lt_DATOSEMPRESA.CurrentIndex = i;
                        datosempresa = new INFO_SOC();
                        datosempresa.BUKRS = lt_DATOSEMPRESA.GetString("BUKRS");
                        datosempresa.BUTXT = lt_DATOSEMPRESA.GetString("BUTXT");
                        datosempresa.STCD1 = lt_DATOSEMPRESA.GetString("STCD1");
                        DatosEmpresa.Add(datosempresa);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("{0} Exception caught.", ex);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                }
                    try
                    {
                        lt_DATOS_VP = BapiGetUser.GetTable("DATOS_VP");
                        for (int i = 0; i < lt_DATOS_VP.Count(); i++)
                        {
                            lt_DATOS_VP.CurrentIndex = i;
                            datosvp = new DATOS_VP();
                            datosvp.NUM_POS = lt_DATOS_VP.GetString("NUM_POS");
                            datosvp.DESCRIP_VP = lt_DATOS_VP.GetString("DESCRIP_VP");
                            datosvp.NUM_VP = lt_DATOS_VP.GetString("NUM_VP");
                            datosvp.FECHA_EMISION = lt_DATOS_VP.GetString("FECHA_EMISION");
                            datosvp.FECHA_VENC = lt_DATOS_VP.GetString("FECHA_VENC");
                            datosvp.MONEDA = lt_DATOS_VP.GetString("MONEDA");
                            string str = datosvp.MONEDA;
                            switch (str)
                            {
                               
                                case "USD":
                                            //string Valor = lt_DATOS_VP.GetString("MONTO_MO").Trim();
                                            //Valor = Valor.Replace(".", "");
                                            //Valor = Valor.Replace(",", "");
                                            //decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                            //string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                            datosvp.MONTO_MO = Formato.FormatoMonedaExtranjera(lt_DATOS_VP.GetString("MONTO_MO").Trim());

                                            //string Valor2 = lt_DATOS_VP.GetString("MONTO_ML").Trim();
                                            //Valor2 = Valor2.Replace(".", "");
                                            //Valor2 = Valor2.Replace(",", "");
                                            //decimal ValorAux2 = Convert.ToDecimal(Valor2.Substring(0, Valor2.Length - 2));
                                            //string Cualquiernombre2 = string.Format("{0:0,0}", ValorAux2);
                                            datosvp.MONTO_ML = Formato.FormatoMonedaExtranjera(lt_DATOS_VP.GetString("MONTO_ML").Trim());
                                    break;
            
                                case "EUR":
                                            //string Valor1 = lt_DATOS_VP.GetString("MONTO_MO").Trim();
                                            //Valor1 = Valor1.Replace(".", "");
                                            //Valor1 = Valor1.Replace(",", "");
                                            //decimal ValorAux1 = Convert.ToDecimal(Valor1.Substring(0, Valor1.Length - 2));
                                            //string Cualquiernombre1 = string.Format("{0:0,0}", ValorAux1);
                                            datosvp.MONTO_MO = Formato.FormatoMonedaExtranjera(lt_DATOS_VP.GetString("MONTO_MO").Trim());

                                            //string Valor5 = lt_DATOS_VP.GetString("MONTO_ML").Trim();
                                            //Valor5 = Valor5.Replace(".", "");
                                            //Valor5 = Valor5.Replace(",", "");
                                            //decimal ValorAux5 = Convert.ToDecimal(Valor5.Substring(0, Valor5.Length - 2));
                                            //string Cualquiernombre5 = string.Format("{0:0,0}", ValorAux5);
                                            datosvp.MONTO_ML = Formato.FormatoMonedaExtranjera(lt_DATOS_VP.GetString("MONTO_ML").Trim());
                                    break;
                                case "CLP":
                                            datosvp.MONTO_MO =  Formato.FormatoMoneda(lt_DATOS_VP.GetString("MONTO_MO"));
                                            datosvp.MONTO_ML =  Formato.FormatoMoneda(lt_DATOS_VP.GetString("MONTO_ML"));
                                    break;
                            }
                            DatosViaPago.Add(datosvp);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("{0} Exception caught.", ex);
                        System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                }
                GC.Collect();
            }

            catch (InvalidCastException ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }
    }
}

