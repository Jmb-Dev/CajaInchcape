using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using SAP.Middleware.Connector;
using System.Configuration;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.CierreCaja.Estructura;

namespace CajaIndu.AppPersistencia.Class.CierreCaja
{
    class CierreCaja
    {
        public string errormessage = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public int diasatraso = 0;
        ConexSAP connectorSap = new ConexSAP();
        public List<ESTATUS> T_Retorno = new List<ESTATUS>();

        public void cierrecaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_ID_CAJA, string P_LAND, string P_MONTO_CIERRE, string P_MONTO_DIF, string P_COMENTARIO_DIF, string P_COMENTARIO_CIERRE)
        {
            try
            {
                T_Retorno.Clear();
                errormessage = "";
                status = "";
                message = "";
                stringRfc = "";
                IRfcTable lt_CIERRE_CAJA_DET_EFECT;
                ESTATUS retorno;
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_FM_REG_CIERRE");

                    BapiGetUser.SetValue("ID_CAJA", P_ID_CAJA);
                    BapiGetUser.SetValue("USUARIO", P_UNAME);
                    BapiGetUser.SetValue("LAND", P_LAND);
                    BapiGetUser.SetValue("MONTO_CIERRE", P_MONTO_CIERRE);
                    BapiGetUser.SetValue("MONTO_DIF", P_MONTO_DIF);
                    BapiGetUser.SetValue("COMENTARIO_DIF", P_COMENTARIO_DIF);
                    BapiGetUser.SetValue("COMENTARIO_CIERRE", P_COMENTARIO_CIERRE);

                    BapiGetUser.Invoke(SapRfcDestination);
                    //LLenamos los datos que retorna la estructura de la RFC
                    //lt_CIERRE_CAJA = BapiGetUser.GetTable("ESTATUS");
                    diasatraso = BapiGetUser.GetInt("DIAS_ATRASO");

                    lt_CIERRE_CAJA_DET_EFECT = BapiGetUser.GetTable("ESTATUS");
                    for (int i = 0; i < lt_CIERRE_CAJA_DET_EFECT.Count(); i++)
                    {
                        lt_CIERRE_CAJA_DET_EFECT.CurrentIndex = i;
                        retorno = new ESTATUS();

                        retorno.TYPE = lt_CIERRE_CAJA_DET_EFECT.GetString("TYPE");
                        if (i==0)
                            status = lt_CIERRE_CAJA_DET_EFECT.GetString("TYPE");
                        retorno.ID = lt_CIERRE_CAJA_DET_EFECT.GetString("ID");
                        retorno.NUMBER = lt_CIERRE_CAJA_DET_EFECT.GetString("NUMBER");
                        retorno.MESSAGE = lt_CIERRE_CAJA_DET_EFECT.GetString("MESSAGE");
                        retorno.LOG_NO = lt_CIERRE_CAJA_DET_EFECT.GetString("LOG_NO");
                        retorno.LOG_MSG_NO = lt_CIERRE_CAJA_DET_EFECT.GetString("LOG_MSG_NO");
                        retorno.MESSAGE_V1 = lt_CIERRE_CAJA_DET_EFECT.GetString("MESSAGE_V1");
                        retorno.MESSAGE_V2 = lt_CIERRE_CAJA_DET_EFECT.GetString("MESSAGE_V2");
                        retorno.MESSAGE_V3 = lt_CIERRE_CAJA_DET_EFECT.GetString("MESSAGE_V3");
                        retorno.MESSAGE_V4 = lt_CIERRE_CAJA_DET_EFECT.GetString("MESSAGE_V4");
                        retorno.PARAMETER = lt_CIERRE_CAJA_DET_EFECT.GetString("PARAMETER");
                        retorno.ROW = lt_CIERRE_CAJA_DET_EFECT.GetString("ROW");
                        retorno.FIELD = lt_CIERRE_CAJA_DET_EFECT.GetString("FIELD");
                        retorno.SYSTEM = lt_CIERRE_CAJA_DET_EFECT.GetString("SYSTEM");
                        T_Retorno.Add(retorno);
                    }

                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
            }
           //return T_Retorno;
            GC.Collect();
        }
    }
}
                  
//                    CERR_CAJA_resp = new CERR_CAJA();
//                    Convert.ToString(BapiGetUser.GetValue("ESTATUS"));
//                    //CERR_CAJA_resp.ESTATUS = lt_CIERRE_CAJA.GetStructure("DET_EFECTIVO").GetString("MESSAGE");
//                    //for (int i = 0; i <= lt_CIERRE_CAJA.Count(); i++)
//                    //{
//                    //    //lt_CIERRE_CAJA.CurrentIndex = i;
                        

//                    //    CERR_CAJA_resp.DIAS_ATRASO = lt_CIERRE_CAJA[i].GetString("DIAS_ATRASO");
//                    //    CERR_CAJA_resp.ESTATUS = lt_CIERRE_CAJA[i].GetString("ESTATUS");


//                    //    T_Retorno.Add(CERR_CAJA_resp);

//                    //}

//                }
//                else
//                ////Si el valor de retorno es distinto nulo o vacio,se emite el mensaje de error.
//                {
//                    errormessage = retval;
//                }
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine(ex.Message + ex.StackTrace);
//System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
//            }
//        }
//    }

//}
