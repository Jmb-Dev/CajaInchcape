using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.GestionDeDepositos.Estructura;
using CajaIndigo.AppPersistencia.Class.DepositoProceso.Estructura;

namespace CajaIndigo.AppPersistencia.Class.DepositoProceso
{
    class DepositoProceso
    {
        public string errormessage = "";
        public string message = "";
        public string IdCaja = "";
        public string Efectivo = "";
        public string NumComprob = "";
        public string comprobante = "";
        public string Deposito = "";

        public List<VIAS_PAGOGD> vpgestiondepositos = new List<VIAS_PAGOGD>();
        public List<CajaIndigo.AppPersistencia.Class.DepositoProceso.Estructura.RETORNO> Retorno = new List<CajaIndigo.AppPersistencia.Class.DepositoProceso.Estructura.RETORNO>();
        //public List<BCO_DESTINO> BancoDest = new List<BCO_DESTINO>();
        ConexSAP connectorSap = new ConexSAP();

        public void depositoproceso(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE
            , string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_USUARIO, string P_PAIS, string P_ID_APERTURA
            , string P_ID_CIERRE, string P_ID_ARQUEO,string P_FECHADEPOS, string P_NUMDEPOS, List<VIAS_PAGOGD> P_VIASPAGO, string P_DATOSBANCO, string P_HKONT)
        {
            try
            {
                CajaIndigo.AppPersistencia.Class.DepositoProceso.Estructura.RETORNO retorno;
                VIAS_PAGOGD vpgestion;
                
                //DETALLE_REND detallerend;
                vpgestiondepositos.Clear();
                Retorno.Clear();
                errormessage = "";
                message = "";
                IdCaja = "";
                Efectivo = "0";

                IRfcTable lt_RETORNO;
                IRfcTable lt_VPGESTIONBANCOS;
                

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_PROC_DEPOSITOS");

                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_USUARIO);
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);
                    BapiGetUser.SetValue("ID_CIERRE", P_ID_CIERRE);
                    BapiGetUser.SetValue("ID_ARQUEO", P_ID_ARQUEO);

                    IRfcTable GralDat2 = BapiGetUser.GetTable("VIAS_PAGO_IN");
                    try
                    {
                        for (var i = 0; i < P_VIASPAGO.Count; i++)
                        {
                            GralDat2.Append();
                            GralDat2.SetValue("SELECCION", "X");
                            GralDat2.SetValue("ID_CAJA", P_VIASPAGO[i].ID_CAJA);
                            GralDat2.SetValue("ID_CIERRE", P_ID_CIERRE);
                            GralDat2.SetValue("TEXT_VIA_PAGO", P_VIASPAGO[i].TEXT_VIA_PAGO);
                            GralDat2.SetValue("FECHA_EMISION", P_VIASPAGO[i].FECHA_EMISION);
                            GralDat2.SetValue("NUM_DOC", P_VIASPAGO[i].NUM_DOC);
                            GralDat2.SetValue("TEXT_BANCO", P_VIASPAGO[i].TEXT_BANCO);
                            GralDat2.SetValue("MONTO_DOC", P_VIASPAGO[i].MONTO_DOC);
                            GralDat2.SetValue("ZUONR", P_VIASPAGO[i].ZUONR);
                            GralDat2.SetValue("FECHA_VENC", P_VIASPAGO[i].FECHA_VENC);
                            GralDat2.SetValue("MONEDA", P_VIASPAGO[i].MONEDA);
                            GralDat2.SetValue("ID_BANCO", P_VIASPAGO[i].ID_BANCO);
                            GralDat2.SetValue("VIA_PAGO", P_VIASPAGO[i].VIA_PAGO);
                            GralDat2.SetValue("NUM_DEPOSITO", P_NUMDEPOS);
                            GralDat2.SetValue("USUARIO", P_VIASPAGO[i].USUARIO);
                            GralDat2.SetValue("ID_DEPOSITO", P_VIASPAGO[i].ID_DEPOSITO);
                            GralDat2.SetValue("FEC_DEPOSITO", Convert.ToDateTime( P_FECHADEPOS));
                            //gdepot.BancoDest[i].BANKN + "-" + gdepot.BancoDest[i].BANKL + "-" + gdepot.BancoDest[i].BANKA
                            int Posicion = 0;
                            int PosicionFinal = 0;
                            Posicion = P_DATOSBANCO.IndexOf("-");
                            string CodCuenta = P_DATOSBANCO.Substring(0, Posicion);
                            PosicionFinal = P_DATOSBANCO.LastIndexOf("-");
                            string CodBanco = P_DATOSBANCO.Substring(Posicion + 1, (PosicionFinal)-(Posicion+1));
                            //GralDat2.SetValue("BANCO", P_VIASPAGO[i].BANCO);
                            GralDat2.SetValue("BANCO", CodBanco);

                            GralDat2.SetValue("CTA_BANCO",CodCuenta );
                            GralDat2.SetValue("BELNR_DEP", P_VIASPAGO[i].BELNR_DEP);
                            GralDat2.SetValue("BELNR", P_VIASPAGO[i].BELNR);
                            GralDat2.SetValue("SOCIEDAD", P_VIASPAGO[i].SOCIEDAD);
                            GralDat2.SetValue("HKONT", P_HKONT);
                            GralDat2.SetValue("ID_COMPROBANTE", P_VIASPAGO[i].ID_COMPROBANTE);
                            GralDat2.SetValue("ID_DETALLE", P_VIASPAGO[i].ID_DETALLE);
                            
                           
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message + ex.StackTrace);
                    }
                    BapiGetUser.SetValue("VIAS_PAGO_IN", GralDat2);

                    BapiGetUser.Invoke(SapRfcDestination);

                    Deposito = BapiGetUser.GetString("ID_DEPOSITO");
                    lt_VPGESTIONBANCOS = BapiGetUser.GetTable("VIAS_PAGO_IN");

                    if (lt_VPGESTIONBANCOS.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_VPGESTIONBANCOS.RowCount; i++)
                        {
                            try
                            {
                                lt_VPGESTIONBANCOS.CurrentIndex = i;
                                vpgestion = new VIAS_PAGOGD();

                                vpgestion.SELECCION = lt_VPGESTIONBANCOS[i].GetString("SELECCION");
                                vpgestion.ID_CAJA = lt_VPGESTIONBANCOS[i].GetString("ID_CAJA");
                                vpgestion.ID_APERTURA = lt_VPGESTIONBANCOS[i].GetString("ID_APERTURA");
                                vpgestion.ID_CIERRE = lt_VPGESTIONBANCOS[i].GetString("ID_CIERRE");
                                vpgestion.TEXT_VIA_PAGO = lt_VPGESTIONBANCOS[i].GetString("TEXT_VIA_PAGO");
                                vpgestion.FECHA_EMISION = lt_VPGESTIONBANCOS[i].GetString("FECHA_EMISION");
                                vpgestion.NUM_DOC = lt_VPGESTIONBANCOS[i].GetString("NUM_DOC");
                                vpgestion.TEXT_BANCO = lt_VPGESTIONBANCOS[i].GetString("TEXT_BANCO");
                                vpgestion.MONTO_DOC = lt_VPGESTIONBANCOS[i].GetString("MONTO_DOC");
                                vpgestion.ZUONR = lt_VPGESTIONBANCOS[i].GetString("ZUONR");
                                vpgestion.FECHA_VENC = lt_VPGESTIONBANCOS[i].GetString("FECHA_VENC");
                                vpgestion.MONEDA = lt_VPGESTIONBANCOS[i].GetString("MONEDA");
                                vpgestion.ID_BANCO = lt_VPGESTIONBANCOS[i].GetString("ID_BANCO");
                                vpgestion.VIA_PAGO = lt_VPGESTIONBANCOS[i].GetString("VIA_PAGO");
                                vpgestion.SOCIEDAD = lt_VPGESTIONBANCOS[i].GetString("SOCIEDAD");
                                vpgestion.NUM_DEPOSITO = lt_VPGESTIONBANCOS[i].GetString("NUM_DEPOSITO");
                                vpgestion.USUARIO = lt_VPGESTIONBANCOS[i].GetString("USUARIO");
                                vpgestion.ID_DEPOSITO = lt_VPGESTIONBANCOS[i].GetString("ID_DEPOSITO");
                                vpgestion.FEC_DEPOSITO = lt_VPGESTIONBANCOS[i].GetString("FEC_DEPOSITO");
                                vpgestion.BANCO = lt_VPGESTIONBANCOS[i].GetString("BANCO");
                                vpgestion.CTA_BANCO = lt_VPGESTIONBANCOS[i].GetString("CTA_BANCO");
                                vpgestion.BELNR_DEP = lt_VPGESTIONBANCOS[i].GetString("BELNR_DEP");
                                vpgestion.BELNR = lt_VPGESTIONBANCOS[i].GetString("BELNR");
                                vpgestion.SOCIEDAD = lt_VPGESTIONBANCOS[i].GetString("SOCIEDAD");
                                vpgestion.ID_COMPROBANTE = lt_VPGESTIONBANCOS[i].GetString("ID_COMPROBANTE");
                                vpgestion.ID_DETALLE = lt_VPGESTIONBANCOS[i].GetString("ID_DETALLE");
                                vpgestiondepositos.Add(vpgestion);
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
                        System.Windows.Forms.MessageBox.Show("No existe(n) registro(s)");
                    }

                    lt_RETORNO = BapiGetUser.GetTable("RETORNO");
                    try
                    {
                        for (int i = 0; i < lt_RETORNO.Count(); i++)
                        {
                            lt_RETORNO.CurrentIndex = i;
                            retorno = new CajaIndigo.AppPersistencia.Class.DepositoProceso.Estructura.RETORNO();
                            if (lt_RETORNO.GetString("TYPE") == "S")
                            {
                                message = message + " - " + lt_RETORNO.GetString("MESSAGE") + "-" + lt_RETORNO.GetString("MESSAGE_V1").Trim();
                                NumComprob = lt_RETORNO.GetString("MESSAGE_V4").Trim(); ;
                            }
                            if (lt_RETORNO.GetString("TYPE") == "E")
                            {
                                errormessage = errormessage + " - " + lt_RETORNO.GetString("MESSAGE");
                            }
                            retorno.TYPE = lt_RETORNO.GetString("TYPE");
                            retorno.CODE = lt_RETORNO.GetString("CODE");
                           // retorno.NUMBER = lt_RETORNO.GetString("NUMBER");
                            retorno.MESSAGE = lt_RETORNO.GetString("MESSAGE");
                            retorno.LOG_NO = lt_RETORNO.GetString("LOG_NO");
                            retorno.LOG_MSG_NO = lt_RETORNO.GetString("LOG_MSG_NO");
                            retorno.MESSAGE_V1 = lt_RETORNO.GetString("MESSAGE_V1");
                            retorno.MESSAGE_V2 = lt_RETORNO.GetString("MESSAGE_V2");
                            retorno.MESSAGE_V3 = lt_RETORNO.GetString("MESSAGE_V3");
                            if (lt_RETORNO.GetString("MESSAGE_V1") != "")
                            {
                                comprobante = comprobante + "-" + lt_RETORNO.GetString("MESSAGE_V1");
                            }
                            retorno.MESSAGE_V4 = lt_RETORNO.GetString("MESSAGE_V4");
                            Retorno.Add(retorno);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message, ex.StackTrace);
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }

                }
                else
                {
                    errormessage = retval;
                    MessageBox.Show("No se pudo conectar a la RFC");
                }
                GC.Collect();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message, ex.StackTrace);
                MessageBox.Show(ex.Message + ex.StackTrace);
            }

        }
    }
}

