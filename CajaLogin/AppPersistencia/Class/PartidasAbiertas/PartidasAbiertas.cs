using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Globalization;
//using System.Windows.Media.Imaging;
//using System.Windows.Shapes;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndigo.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;

namespace CajaIndigo.AppPersistencia.Class.PartidasAbiertas
{
    class PartidasAbiertas
    {
       public List<T_DOCUMENTOS> ObjDatosPartidasOpen = new List<T_DOCUMENTOS>();
       public List<ESTADO> Retorno = new List<ESTADO>();
       public string errormessage = "";
       public string protesto = "";
       ConexSAP connectorSap = new ConexSAP();

       public void partidasopen(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_CODCLIENTE, string P_DOCUMENTO, string P_RUT,
           string P_SOCIEDAD, DateTime P_FECHA_VENC,string P_LAND, string P_FACT_SAP,  string TipoBusqueda)  
        {

            ObjDatosPartidasOpen.Clear();
            Retorno.Clear();
            IRfcTable lt_t_documentos;
            IRfcStructure lt_retorno;

          //  PART_ABIERTAS  PART_ABIERTAS_resp;
             T_DOCUMENTOS PART_ABIERTAS_resp;
             ESTADO retorno_resp;
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_DOC");
                    BapiGetUser.SetValue("CODCLIENTE", P_CODCLIENTE);
                    BapiGetUser.SetValue("DOCUMENTO", P_DOCUMENTO);
                    BapiGetUser.SetValue("LAND", P_LAND);
                    BapiGetUser.SetValue("RUT", P_RUT);
                    BapiGetUser.SetValue("SOCIEDAD", P_SOCIEDAD);         
                    BapiGetUser.SetValue("FACTURA_SAP", P_DOCUMENTO);
                    BapiGetUser.SetValue("FECHA_VENC", "");

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
                               PART_ABIERTAS_resp = new T_DOCUMENTOS();

                               PART_ABIERTAS_resp.SOCIEDAD = lt_t_documentos[i].GetString("SOCIEDAD");
                               PART_ABIERTAS_resp.NDOCTO = lt_t_documentos[i].GetString("NDOCTO");
                               PART_ABIERTAS_resp.NREF = lt_t_documentos[i].GetString("NREF");
                               PART_ABIERTAS_resp.CLASE_CUENTA = lt_t_documentos[i].GetString("CLASE_CUENTA");
                               PART_ABIERTAS_resp.CLASE_DOC = lt_t_documentos[i].GetString("CLASE_DOC");
                               PART_ABIERTAS_resp.COD_CLIENTE = lt_t_documentos[i].GetString("COD_CLIENTE");
                               PART_ABIERTAS_resp.RUTCLI = lt_t_documentos[i].GetString("RUTCLI");
                               PART_ABIERTAS_resp.NOMCLI = lt_t_documentos[i].GetString("NOMCLI");
                               PART_ABIERTAS_resp.CEBE = lt_t_documentos[i].GetString("CEBE");
                               DateTime fec_doc = Convert.ToDateTime(lt_t_documentos[i].GetString("FECHA_DOC"));
                               PART_ABIERTAS_resp.FECHA_DOC = fec_doc.ToString("dd/MM/yyyy");
                               DateTime fec_venc = Convert.ToDateTime(lt_t_documentos[i].GetString("FECVENCI"));
                               PART_ABIERTAS_resp.FECVENCI = fec_venc.ToString("dd/MM/yyyy");
                               PART_ABIERTAS_resp.DIAS_ATRASO = lt_t_documentos[i].GetString("DIAS_ATRASO");
                               PART_ABIERTAS_resp.ESTADO = lt_t_documentos[i].GetString("ESTADO");
                               PART_ABIERTAS_resp.ICONO = lt_t_documentos[i].GetString("ICONO");
                               PART_ABIERTAS_resp.MONEDA = lt_t_documentos[i].GetString("MONEDA");
                               PART_ABIERTAS_resp.ACC = lt_t_documentos[i].GetString("ACC");
                               PART_ABIERTAS_resp.CLASE_CUENTA = lt_t_documentos[i].GetString("CLASE_CUENTA");
                               PART_ABIERTAS_resp.COND_PAGO = lt_t_documentos[i].GetString("COND_PAGO");
                               PART_ABIERTAS_resp.CME = lt_t_documentos[i].GetString("CME");
                               PART_ABIERTAS_resp.CONTROL_CREDITO = lt_t_documentos[i].GetString("CONTROL_CREDITO");                             
                               string Monto = "";                            
                                //*******
                               if (lt_t_documentos[i].GetString("MONTOF") == "")
                               {
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
                                       PART_ABIERTAS_resp.MONTOF = Cualquiernombre.Replace(",",".");
                                   }
                                   else
                                   {
                                       string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF"));
                                       decimal ValorAux = Convert.ToDecimal(moneda);
                                       PART_ABIERTAS_resp.MONTOF = string.Format("{0:0,0.##}", ValorAux).Replace(",",".");
                                   }
                               }
                               else
                               {
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
                                       PART_ABIERTAS_resp.MONTOF = Cualquiernombre.Replace(",",".");
                                   }
                                   else
                                   {
                                       string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF"));
                                       decimal ValorAux = Convert.ToDecimal(moneda);
                                       PART_ABIERTAS_resp.MONTOF = string.Format("{0:0,0.##}", ValorAux).Replace(",",".");
                                   }
                               }
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
                                   PART_ABIERTAS_resp.MONTO = Cualquiernombre.Replace(",",".");
                               }
                               else
                               {
                                   string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTO"));
                                   decimal ValorAux = Convert.ToDecimal(moneda);
                                   PART_ABIERTAS_resp.MONTO = string.Format("{0:0,0.##}", ValorAux).Replace(",",".");
                               }
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
                                   PART_ABIERTAS_resp.MONTOF_ABON = Cualquiernombre.Replace(",",".");
                               }
                               else
                               {
                                   string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF_ABON"));
                                   decimal ValorAux = Convert.ToDecimal(moneda);
                                   PART_ABIERTAS_resp.MONTOF_ABON = string.Format("{0:0,0.##}", ValorAux).Replace(",",".");
                               }   
                             //MONTO A PAGAR
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
                                    string monedachil = string.Format("{0:0,0}", ValorAux);
                                    PART_ABIERTAS_resp.MONTOF_PAGAR = monedachil.Replace(",",".");
                                    
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_t_documentos[i].GetString("MONTOF_PAGAR"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    PART_ABIERTAS_resp.MONTOF_PAGAR = string.Format("{0:0,0.##}", ValorAux).Replace(",",".");
                                }
                                 ObjDatosPartidasOpen.Add(PART_ABIERTAS_resp);
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
                           MessageBox.Show("No existe(n) registro(s) para este número de " + TipoBusqueda);
                       }
                   }
                   catch (Exception ex)
                   {
                       Console.WriteLine(ex.Message + ex.StackTrace);
                       System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                      
                   }
                   String Mensaje = "";
                   if (lt_retorno.Count > 0)
                   {
                       string returning = "";
                       retorno_resp = new ESTADO();
                       for (int i = 0; i < lt_retorno.Count(); i++)
                       {
                           // lt_retorno.CurrentIndex = i;
                           if (i == 0)
                           {
                             returning =   retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                           }
                           retorno_resp.TYPE = lt_retorno.GetString("TYPE");
                           retorno_resp.ID = lt_retorno.GetString("ID");
                           retorno_resp.NUMBER = lt_retorno.GetString("NUMBER");
                           retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                           retorno_resp.LOG_NO = lt_retorno.GetString("LOG_NO");
                           retorno_resp.LOG_MSG_NO = lt_retorno.GetString("LOG_MSG_NO");
                           retorno_resp.MESSAGE = lt_retorno.GetString("MESSAGE");
                           retorno_resp.MESSAGE_V1 = lt_retorno.GetString("MESSAGE_V1");
                           if (i == 0)
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
                       if (returning != "")
                       {
                           System.Windows.MessageBox.Show(Mensaje);
                       }
                   }
                   GC.Collect();
                }
                else
                {
                    errormessage = retval;
                    GC.Collect();
                }
            }
    }
    }

