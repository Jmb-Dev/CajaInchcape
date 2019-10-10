using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.GestionDeDepositos.Estructura;

namespace CajaIndigo.AppPersistencia.Class.GestionDeDepositos
{
    class GestionDeDepositos
    {
        public string errormessage = "";
        public string message = "";
        public string IdCaja = "";
        public string Efectivo = "";
        public string NumComprob = "";

        public List<VIAS_PAGOGD> vpgestiondepositos = new List<VIAS_PAGOGD>();
        public List<RETORNO> Retorno = new List<RETORNO>();
        public List<BCO_DESTINO> BancoDest = new List<BCO_DESTINO>();
        public List<BCO_DEPOSITOS> BancoDeposito = new List<BCO_DEPOSITOS>();
        ConexSAP connectorSap = new ConexSAP();

        public void gestiondedepositos(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE
            , string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_USUARIO, string P_PAIS, string P_ID_APERTURA
            , string P_ID_CIERRE, string P_ID_ARQUEO)
            
        {
            try
            {
                RETORNO retorno;
                VIAS_PAGOGD vpgestion;
                BCO_DESTINO banco_destino;
                BCO_DEPOSITOS banco_depositos;
                
                //DETALLE_REND detallerend;
                vpgestiondepositos.Clear();
                Retorno.Clear();
                BancoDest.Clear();
                BancoDeposito.Clear();
                errormessage = "";
                message = "";
                IdCaja = "";
                Efectivo = "0";

                IRfcTable lt_RETORNO;
                IRfcTable lt_VPGESTIONBANCOS;
                IRfcTable lt_BANCODEST;
                IRfcTable lt_BANCODEPOSITOS;

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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_GESTION_DEPOSITOS");

                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_USUARIO);
                    BapiGetUser.SetValue("ID_APERTURA", P_ID_APERTURA);
                    BapiGetUser.SetValue("ID_CIERRE", P_ID_CIERRE);
                    BapiGetUser.SetValue("ID_ARQUEO", P_ID_ARQUEO);
                    
                    BapiGetUser.Invoke(SapRfcDestination);
                    //BapiGetUser.SetValue("I_VBELN",P_NUMDOCSD);
                    //IRfcTable GralDat = BapiGetUser.GetTable("VIAS_PAGO");
                    lt_VPGESTIONBANCOS = BapiGetUser.GetTable("VIAS_PAGO");

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
                                vpgestion.ID_CIERRE = P_ID_CIERRE;
                                vpgestion.TEXT_VIA_PAGO = lt_VPGESTIONBANCOS[i].GetString("TEXT_VIA_PAGO");
                                vpgestion.FECHA_EMISION = lt_VPGESTIONBANCOS[i].GetString("FECHA_EMISION");
                                vpgestion.NUM_DOC = lt_VPGESTIONBANCOS[i].GetString("NUM_DOC");
                                vpgestion.TEXT_BANCO = lt_VPGESTIONBANCOS[i].GetString("TEXT_BANCO");
                                if (lt_VPGESTIONBANCOS[i].GetString("MONEDA") == "CLP")
                                {
                                    string Valor = lt_VPGESTIONBANCOS[i].GetString("MONTO_DOC").Trim();
                                    if (Valor.Contains("-"))
                                    {
                                        Valor = "-" + Valor.Replace("-", "");
                                    }
                                    Valor = Valor.Replace(".", "");
                                    Valor = Valor.Replace(",", "");
                                    decimal ValorAux = Convert.ToDecimal(Valor.Substring(0, Valor.Length - 2));
                                    string monedachile = string.Format("{0:0,0}", ValorAux);
                                    vpgestion.MONTO_DOC = monedachile;                                 
                                }
                                else
                                {
                                    string moneda = Convert.ToString(lt_VPGESTIONBANCOS[i].GetString("MONTO_DOC"));
                                    decimal ValorAux = Convert.ToDecimal(moneda);
                                    vpgestion.MONTO_DOC = string.Format("{0:0,0.##}", ValorAux);
                                   
                                }
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
                                vpgestion.HKONT = lt_VPGESTIONBANCOS[i].GetString("HKONT");
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
                        System.Windows.Forms.MessageBox.Show("No existe(n) registro(s) de vias de pago");
                    }

                    lt_BANCODEST = BapiGetUser.GetTable("BCO_DESTINO");

                    if (lt_BANCODEST.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_BANCODEST.RowCount; i++)
                        {
                            try
                            {
                                lt_BANCODEST.CurrentIndex = i;
                                banco_destino = new BCO_DESTINO();
                                banco_destino.BUKRS = lt_BANCODEST[i].GetString("BUKRS");
                                banco_destino.HBKID = lt_BANCODEST[i].GetString("HBKID");
                                banco_destino.HKTID = lt_BANCODEST[i].GetString("HKTID");
                                banco_destino.BANKN = lt_BANCODEST[i].GetString("BANKN");
                                banco_destino.BANKL = lt_BANCODEST[i].GetString("BANKL");
                                banco_destino.BANKA = lt_BANCODEST[i].GetString("BANKA");
                                banco_destino.WAERS = lt_BANCODEST[i].GetString("WAERS");
                                banco_destino.TEXT1 = lt_BANCODEST[i].GetString("TEXT1");
                               
                                BancoDest.Add(banco_destino);
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
                        System.Windows.Forms.MessageBox.Show("No existe(n) registro(s) en banco destino");
                    }

                    lt_BANCODEPOSITOS = BapiGetUser.GetTable("BCO_DEPOSITOS");

                    if (lt_BANCODEPOSITOS.Count > 0)
                    {
                        //LLenamos la tabla de salida lt_DatGen
                        for (int i = 0; i < lt_BANCODEPOSITOS.RowCount; i++)
                        {
                            try
                            {
                                lt_BANCODEPOSITOS.CurrentIndex = i;
                                banco_depositos = new BCO_DEPOSITOS();
                                banco_depositos.MANDT = lt_BANCODEPOSITOS[i].GetString("MANDT");
                                banco_depositos.BANKS = lt_BANCODEPOSITOS[i].GetString("BANKS");
                                banco_depositos.BUKRS = lt_BANCODEPOSITOS[i].GetString("BUKRS");
                                banco_depositos.WAERS = lt_BANCODEPOSITOS[i].GetString("WAERS");
                                banco_depositos.BANKL = lt_BANCODEPOSITOS[i].GetString("BANKL");
                                banco_depositos.HBKID = lt_BANCODEPOSITOS[i].GetString("HBKID");
                                banco_depositos.BANKN = lt_BANCODEPOSITOS[i].GetString("BANKN");
                                banco_depositos.BANKA = lt_BANCODEPOSITOS[i].GetString("BANKA");
                                banco_depositos.HKONT = lt_BANCODEPOSITOS[i].GetString("HKONT");

                                BancoDeposito.Add(banco_depositos);
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
                        System.Windows.Forms.MessageBox.Show("No existe(n) registro(s) en depósitos de banco");
                    }
                    lt_RETORNO = BapiGetUser.GetTable("RETORNO");
                    try
                    {
                        for (int i = 0; i < lt_RETORNO.Count(); i++)
                        {
                            lt_RETORNO.CurrentIndex = i;
                            retorno = new RETORNO();
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
                            retorno.ID = lt_RETORNO.GetString("ID");
                            retorno.NUMBER = lt_RETORNO.GetString("NUMBER");
                            retorno.MESSAGE = lt_RETORNO.GetString("MESSAGE");
                            retorno.LOG_NO = lt_RETORNO.GetString("LOG_NO");
                            retorno.LOG_MSG_NO = lt_RETORNO.GetString("LOG_MSG_NO");
                            retorno.MESSAGE_V1 = lt_RETORNO.GetString("MESSAGE_V1");
                            retorno.MESSAGE_V2 = lt_RETORNO.GetString("MESSAGE_V2");
                            retorno.MESSAGE_V3 = lt_RETORNO.GetString("MESSAGE_V3");
                            if (lt_RETORNO.GetString("MESSAGE_V4") != "")
                            {
                                // comprobante = ls_RETORNO.GetString("MESSAGE_V4");
                            }
                            retorno.MESSAGE_V4 = lt_RETORNO.GetString("MESSAGE_V4");
                            retorno.PARAMETER = lt_RETORNO.GetString("PARAMETER");
                            retorno.ROW = lt_RETORNO.GetString("ROW");
                            retorno.FIELD = lt_RETORNO.GetString("FIELD");
                            retorno.SYSTEM = lt_RETORNO.GetString("SYSTEM");
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

