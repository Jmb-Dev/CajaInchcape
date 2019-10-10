using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndu.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;
using CajaIndu.AppPersistencia.Class.Connections;

namespace CajaIndu.AppPersistencia.Class.Anticipos
{
    class Anticipos
    {
        public List<T_DOCUMENTOS> ObjDatosAnticipos = new List<T_DOCUMENTOS>();
        public List<ESTADO> Retorno = new List<ESTADO>();
        public string errormessage = "";
        public string protesto = "";
        ConexSAP connectorSap = new ConexSAP();

        public void anticiposopen(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA
            , string P_DOCUMENTO, string P_RUT, string P_SOCIEDAD, string P_LAND ,  string TipoBusqueda)  

        {
            ObjDatosAnticipos.Clear();
            Retorno.Clear();
            protesto = "";
            errormessage = "";
            IRfcTable lt_t_documentos;
            IRfcStructure lt_retorno;

            //  PART_ABIERTAS  PART_ABIERTAS_resp;
            T_DOCUMENTOS ANTICIPOS_resp;
            ESTADO retorno_resp; 

            //Conexion a SAP
            //connectorSap.idioma = "ES";
            //connectorSap.idSistema = "INS";
            //connectorSap.instancia = "00";
            //connectorSap.mandante = "400";
            //connectorSap.paswr = P_PASSWORD;
            //connectorSap.sapRouter = "/H/64.76.139.78/H/";
            //connectorSap.user = P_UNAME;
            //connectorSap.server = "10.9.100.4";
            //frm.txtIdSistema.Text = txtIdSistema.Text;
            //frm.txtInstancia.Text = txtInstancia.Text;
            //frm.txtMandante.Text = txtMandante.Text;
            //frm.txtSapRouter.Text = txtSapRouter.Text;
            //frm.txtServer.Text = txtServer.Text;
            //frm.txtIdioma.Text = txtIdioma.Text;
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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_ANT");
                BapiGetUser.SetValue("DOCUMENTO", P_DOCUMENTO);
                BapiGetUser.SetValue("LAND", P_LAND);
                BapiGetUser.SetValue("RUT", P_RUT);
                BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);
                //BapiGetUser.SetValue("PROT", P_PROTESTO);
                BapiGetUser.Invoke(SapRfcDestination);
                protesto = BapiGetUser.GetString("PE_PROTESTADO");

                lt_t_documentos = BapiGetUser.GetTable("T_DOCUMENTOS");
                lt_retorno = BapiGetUser.GetStructure("SE_ESTATUS");
                //lt_PART_ABIERTAS = BapiGetUser.GetTable("ZCLSP_TT_LISTA_DOCUMENTOS");
                try
                {
                    if (lt_t_documentos.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_t_documentos.RowCount; i++)
                        {
                            try
                            {
                                lt_t_documentos.CurrentIndex = i;
                                ANTICIPOS_resp = new T_DOCUMENTOS();

                                ANTICIPOS_resp.NDOCTO = lt_t_documentos[i].GetString("NDOCTO");
                                string Monto = "";
                                int indice = 0;
                                //*******
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTOF").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    ANTICIPOS_resp.MONTOF = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    ANTICIPOS_resp.MONTOF = string.Format("{0:0,0.##}", ValorAux);
                                }

                                //if (lt_t_documentos[i].GetString("MONTOF") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO").Substring(0, indice - 1);
                                //    ANTICIPOS_resp.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    ANTICIPOS_resp.MONTOF = lt_t_documentos[i].GetString("MONTOF");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTO").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    ANTICIPOS_resp.MONTO = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTO"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    ANTICIPOS_resp.MONTO = string.Format("{0:0,0.##}", ValorAux);
                                }

                                //if (lt_t_documentos[i].GetString("MONTO") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO").Substring(0, indice - 1);
                                //    ANTICIPOS_resp.MONTO = Monto;
                                //}
                                //else
                                //{
                                //    ANTICIPOS_resp.MONTO = lt_t_documentos[i].GetString("MONTO");
                                //}
                                ANTICIPOS_resp.MONEDA = lt_t_documentos[i].GetString("MONEDA");
                                ANTICIPOS_resp.FECVENCI = lt_t_documentos[i].GetString("FECVENCI");
                                ANTICIPOS_resp.CONTROL_CREDITO = lt_t_documentos[i].GetString("CONTROL_CREDITO");
                                ANTICIPOS_resp.CEBE = lt_t_documentos[i].GetString("CEBE");
                                ANTICIPOS_resp.COND_PAGO = lt_t_documentos[i].GetString("COND_PAGO");
                                ANTICIPOS_resp.RUTCLI = lt_t_documentos[i].GetString("RUTCLI");
                                ANTICIPOS_resp.NOMCLI = lt_t_documentos[i].GetString("NOMCLI");
                                ANTICIPOS_resp.ESTADO = lt_t_documentos[i].GetString("ESTADO");
                                ANTICIPOS_resp.ICONO = lt_t_documentos[i].GetString("ICONO");
                                ANTICIPOS_resp.DIAS_ATRASO = lt_t_documentos[i].GetString("DIAS_ATRASO");
                                //if (lt_t_documentos[i].GetString("MONTOF_ABON") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO_ABONADO").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO_ABONADO").Substring(0, indice - 1);
                                //    ANTICIPOS_resp.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    ANTICIPOS_resp.MONTOF_ABON = lt_t_documentos[i].GetString("MONTOF_ABON");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTOF_ABON").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    ANTICIPOS_resp.MONTOF_ABON = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF_ABON"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    ANTICIPOS_resp.MONTOF_ABON = string.Format("{0:0,0.##}", ValorAux);
                                }

                                //if (lt_t_documentos[i].GetString("MONTOF_PAGAR") == "")
                                //{
                                //    indice = lt_t_documentos[i].GetString("MONTO_PAGAR").IndexOf(',');
                                //    Monto = lt_t_documentos[i].GetString("MONTO_PAGAR").Substring(0, indice - 1);
                                //    ANTICIPOS_resp.MONTOF = Monto;
                                //}
                                //else
                                //{
                                //    ANTICIPOS_resp.MONTOF_PAGAR = lt_t_documentos[i].GetString("MONTOF_PAGAR");
                                //}
                                if (lt_t_documentos[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_t_documentos[i].GetString("MONTOF_PAGAR").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor);
                                    string Cualquiernombre = string.Format("{0:0,0}", ValorAux);
                                    ANTICIPOS_resp.MONTOF_PAGAR = Cualquiernombre;
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF_PAGAR"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    ANTICIPOS_resp.MONTOF_PAGAR = string.Format("{0:0,0.##}", ValorAux);
                                }
                                ANTICIPOS_resp.NREF = lt_t_documentos[i].GetString("NREF");
                                ANTICIPOS_resp.FECHA_DOC = lt_t_documentos[i].GetString("FECHA_DOC");
                                ANTICIPOS_resp.COD_CLIENTE = lt_t_documentos[i].GetString("COD_CLIENTE");
                                ANTICIPOS_resp.SOCIEDAD = lt_t_documentos[i].GetString("SOCIEDAD");
                                ANTICIPOS_resp.CLASE_DOC = lt_t_documentos[i].GetString("CLASE_DOC");
                                ANTICIPOS_resp.CLASE_CUENTA = lt_t_documentos[i].GetString("CLASE_CUENTA");
                                ANTICIPOS_resp.CME = lt_t_documentos[i].GetString("CME");
                                ANTICIPOS_resp.ACC = lt_t_documentos[i].GetString("ACC");
                                ObjDatosAnticipos.Add(ANTICIPOS_resp);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message + ex.StackTrace);
                                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                            }
                         }
                    }
                    else
                    {
                        MessageBox.Show("No existe(n) registro(s)" );
                    }

                    String Mensaje = "";
                    if (lt_retorno.Count > 0)
                    {
                        retorno_resp = new ESTADO();
                        for (int i = 0; i < lt_retorno.Count(); i++)
                        {
                            // lt_retorno.CurrentIndex = i;

                            retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                            retorno_resp.ID = lt_retorno.GetString("ID");
                            retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                            retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                            retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                            retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                            retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                            retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                            if (lt_retorno.GetString("TYPE") == "S")
                            {
                                Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE") + " - " + lt_retorno.GetString("MESSAGE_V1");
                            }
                            retorno_resp.MESSAGE_V2 = lt_retorno.GetString("MESSAGE_V2");
                            retorno_resp.MESSAGE_V3 = lt_retorno.GetString("MESSAGE_V3");
                            retorno_resp.MESSAGE_V4 = lt_retorno.GetString("MESSAGE_V4");
                            retorno_resp.PARAMETER = lt_retorno.GetString("PARAMETER");
                            retorno_resp.ROW = lt_retorno.GetString("ROW");
                            retorno_resp.FIELD = lt_retorno.GetString("FIELD");
                            retorno_resp.SYSTEM = lt_retorno.GetString("SYSTEM");
                            Retorno.Add(retorno_resp);
                        }
                        //System.Windows.MessageBox.Show(Mensaje);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                    System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);

                }

            }
            else
            {
                errormessage = retval;
            }
            GC.Collect();
        }

    }
}
