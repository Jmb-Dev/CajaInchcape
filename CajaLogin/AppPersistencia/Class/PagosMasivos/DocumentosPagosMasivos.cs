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
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndigo.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;


namespace CajaIndigo.AppPersistencia.Class.DocumentosPagosMasivos
{
    class DocumentosPagosMasivos
    {
        public List<T_DOCUMENTOS> ObjDatosPartidasOpen = new List<T_DOCUMENTOS>(); 
       public string errormessage = "";
       public string protesto = "";
       ConexSAP connectorSap = new ConexSAP();
       public List<ESTADO> Retorno = new List<ESTADO>();

       public void pagosmasivos(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_RUT, string P_SOCIEDAD, List<PagosMasivos> ListaExc)  
        {

            ObjDatosPartidasOpen.Clear();
            Retorno.Clear();
            errormessage = "";
            protesto = "";
            IRfcTable lt_PAGOS_MASIVOS;
            IRfcStructure lt_retorno;

            ESTADO retorno_resp;  
            T_DOCUMENTOS PAGOS_MASIVOS_resp;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GET_DOC_MASI");
                    BapiGetUser.SetValue("STCD1", P_RUT);
                    BapiGetUser.SetValue("BUKRS", P_SOCIEDAD);
                   
                 

                    IRfcTable GralDat = BapiGetUser.GetTable("T_GET_DOC");

                    for (var i = 0; i < ListaExc.Count; i++)
                    {
                        GralDat.Append();
                        GralDat.SetValue("XBLNR", ListaExc[i].Referencia);
                        GralDat.SetValue("MONTO", ListaExc[i].Monto);
                        GralDat.SetValue("WAERS", ListaExc[i].Moneda);
                    }

                    BapiGetUser.SetValue("T_GET_DOC", GralDat);


                    BapiGetUser.Invoke(SapRfcDestination);


                   protesto = BapiGetUser.GetString("PE_PROTESTADO");
                   lt_PAGOS_MASIVOS = BapiGetUser.GetTable("T_DOCUMENTOS");
                   lt_retorno = BapiGetUser.GetStructure("SE_ESTATUS");
                    //lt_PART_ABIERTAS = BapiGetUser.GetTable("ZCLSP_TT_LISTA_DOCUMENTOS");
                   if (lt_PAGOS_MASIVOS.Count > 0)
                   {
                       //LLenamos la tabla de salida lt_DatGen
                       for (int i = 0; i < lt_PAGOS_MASIVOS.RowCount; i++)
                       {
                           lt_PAGOS_MASIVOS.CurrentIndex = i;
                           PAGOS_MASIVOS_resp = new T_DOCUMENTOS();

                           PAGOS_MASIVOS_resp.SOCIEDAD = lt_PAGOS_MASIVOS[i].GetString("SOCIEDAD");
                           PAGOS_MASIVOS_resp.NDOCTO = lt_PAGOS_MASIVOS[i].GetString("NDOCTO");
                           PAGOS_MASIVOS_resp.NREF = lt_PAGOS_MASIVOS[i].GetString("NREF");
                           PAGOS_MASIVOS_resp.CLASE_CUENTA = lt_PAGOS_MASIVOS[i].GetString("CLASE_CUENTA");
                           PAGOS_MASIVOS_resp.CLASE_DOC = lt_PAGOS_MASIVOS[i].GetString("CLASE_DOC");
                           PAGOS_MASIVOS_resp.COD_CLIENTE = lt_PAGOS_MASIVOS[i].GetString("COD_CLIENTE");
                           PAGOS_MASIVOS_resp.RUTCLI = lt_PAGOS_MASIVOS[i].GetString("RUTCLI");
                           PAGOS_MASIVOS_resp.NOMCLI = lt_PAGOS_MASIVOS[i].GetString("NOMCLI");
                           PAGOS_MASIVOS_resp.CEBE = lt_PAGOS_MASIVOS[i].GetString("CEBE");
                           PAGOS_MASIVOS_resp.FECHA_DOC = lt_PAGOS_MASIVOS[i].GetString("FECHA_DOC");
                           PAGOS_MASIVOS_resp.FECVENCI = lt_PAGOS_MASIVOS[i].GetString("FECVENCI");
                           PAGOS_MASIVOS_resp.DIAS_ATRASO = lt_PAGOS_MASIVOS[i].GetString("DIAS_ATRASO");
                           PAGOS_MASIVOS_resp.ESTADO = lt_PAGOS_MASIVOS[i].GetString("ESTADO");
                           PAGOS_MASIVOS_resp.ICONO = lt_PAGOS_MASIVOS[i].GetString("ICONO");
                           PAGOS_MASIVOS_resp.MONEDA = lt_PAGOS_MASIVOS[i].GetString("MONEDA");
                           PAGOS_MASIVOS_resp.ACC = lt_PAGOS_MASIVOS[i].GetString("ACC");
                           PAGOS_MASIVOS_resp.CLASE_CUENTA = lt_PAGOS_MASIVOS[i].GetString("CLASE_CUENTA");
                           PAGOS_MASIVOS_resp.CLASE_DOC = lt_PAGOS_MASIVOS[i].GetString("CLASE_DOC");
                           PAGOS_MASIVOS_resp.CME = lt_PAGOS_MASIVOS[i].GetString("CME");
                           PAGOS_MASIVOS_resp.CONTROL_CREDITO = lt_PAGOS_MASIVOS[i].GetString("CONTROL_CREDITO");
                           //string.Format("{0:0.##}", lvatend)
                           //decimal lvNetoAbo2 = Convert.ToDecimal(t_REPORTE_CONTACTOS[i].GetString("NETO_ABONO2"));
                           //REPORTE_CONTACTOS_resp.NETO_ABONO2 = string.Format("{0:#,0}", lvNetoAbo2);
                           if (lt_PAGOS_MASIVOS[i].GetString("MONTOF") == "")
                               {
                                   PAGOS_MASIVOS_resp.MONTOF = "0";
                               }
                           else
                               {
                                    decimal  Cualquiernombre =Convert.ToDecimal(lt_PAGOS_MASIVOS[i].GetString("MONTOF"));
                                   // PAGOS_MASIVOS_resp.MONTOF = lt_PAGOS_MASIVOS[i].GetString("MONTOF");
                                    PAGOS_MASIVOS_resp.MONTOF = string.Format("{0:0.##}", Cualquiernombre);
                               }
                           if (lt_PAGOS_MASIVOS[i].GetString("MONTO") == "")
                           {
                               PAGOS_MASIVOS_resp.MONTO = "0";
                           }
                           else
                           {
                               decimal Cualquiernombre = Convert.ToDecimal(lt_PAGOS_MASIVOS[i].GetString("MONTO"));
                               // PAGOS_MASIVOS_resp.MONTOF = lt_PAGOS_MASIVOS[i].GetString("MONTOF");
                               PAGOS_MASIVOS_resp.MONTO = string.Format("{0:0.##}", Cualquiernombre);
                           }
                           if (lt_PAGOS_MASIVOS[i].GetString("MONTOF_ABON") == "")
                               {
                                   PAGOS_MASIVOS_resp.MONTOF_ABON = "0";
                               }
                           else
                               {
                                  decimal Cualquiernombre = Convert.ToDecimal(lt_PAGOS_MASIVOS[i].GetString("MONTOF_ABON"));
                                  PAGOS_MASIVOS_resp.MONTOF_ABON =  string.Format("{0:0.##}", Cualquiernombre);
                               }
                           if (lt_PAGOS_MASIVOS[i].GetString("MONTOF_PAGAR") == "")
                               {
                                  PAGOS_MASIVOS_resp.MONTOF_PAGAR = "0";
                               }
                           else
                               {
                                  // PAGOS_MASIVOS_resp.MONTOF_ABON = lt_PAGOS_MASIVOS[i].GetString("MONTOF_ABON");
                                  decimal Cualquiernombre = Convert.ToDecimal(lt_PAGOS_MASIVOS[i].GetString("MONTOF_PAGAR"));
                                  PAGOS_MASIVOS_resp.MONTOF_PAGAR = string.Format("{0:0.##}", Cualquiernombre);
                               }
                         // PAGOS_MASIVOS_resp.MONTOF_PAGAR = lt_PAGOS_MASIVOS[i].GetString("MONTOF_PAGAR");
                        

                           ObjDatosPartidasOpen.Add(PAGOS_MASIVOS_resp);


                       }
                   }
                   else
                   {
                       MessageBox.Show("No existen registros para este número de RUT");
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
                       System.Windows.MessageBox.Show(Mensaje);
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

