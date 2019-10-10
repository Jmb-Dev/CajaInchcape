using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndigo.AppPersistencia.Class.Connections;
using CajaIndigo.AppPersistencia.Class.ReimpresionFiscal.Estructura;

namespace CajaIndigo.AppPersistencia.Class.ReimpresionFiscal
{
    class ReimpresionFiscal
    {
        ConexSAP connectorSap = new ConexSAP();
        string errormessage = "";
        public string url = "";

        public List<DTE_SII> reimprFiscal2 = new List<DTE_SII>();

        public void ReipresionFiscal2(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_DOCUMENTO, string P_SOCIEDAD)
        {
            IRfcTable lt_DTE_SII;
            DTE_SII DTE_SII_resp;
       
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
                try
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_REIMPRESION");
                    BapiGetUser.SetValue("XBLNR", P_DOCUMENTO);
                    BapiGetUser.SetValue("BUKRS", P_SOCIEDAD);


                    BapiGetUser.Invoke(SapRfcDestination);


                    lt_DTE_SII = BapiGetUser.GetTable("DTE_SII");

                    for (int i = 0; i < lt_DTE_SII.RowCount; i++)
                    {

                        lt_DTE_SII.CurrentIndex = i;
                        DTE_SII_resp = new DTE_SII();
                        DTE_SII_resp.VBELN = lt_DTE_SII[i].GetString("VBELN");
                        DTE_SII_resp.BUKRS = lt_DTE_SII[i].GetString("BUKRS");
                        DTE_SII_resp.FECIMP = lt_DTE_SII[i].GetString("FECIMP");
                        DTE_SII_resp.FODOC = lt_DTE_SII[i].GetString("FODOC");
                        DTE_SII_resp.HORIM = lt_DTE_SII[i].GetString("HORIM");
                        DTE_SII_resp.KONDA = lt_DTE_SII[i].GetString("KONDA");
                        DTE_SII_resp.TDSII = lt_DTE_SII[i].GetString("TDSII");
                        DTE_SII_resp.URLSII = lt_DTE_SII[i].GetString("URLSII");
                        DTE_SII_resp.WAERS = lt_DTE_SII[i].GetString("WAERS");
                        DTE_SII_resp.XBLNR = lt_DTE_SII[i].GetString("XBLNR");
                        DTE_SII_resp.ZUONR = lt_DTE_SII[i].GetString("ZUONR");
                        reimprFiscal2.Add(DTE_SII_resp);
                    }
                    //url = BapiGetUser.GetString("URL");
                    //GC.Collect();
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
                GC.Collect();
            }
        
        
        }
        public void reimpresionfiscal(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_DOCUMENTO)
        {

            //IRfcTable lt_t_documentos;
            //IRfcStructure lt_retorno;

            //  PART_ABIERTAS  PART_ABIERTAS_resp;


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
                try
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(connectorSap.connectorConfig);
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_REIMPRESION");
                    BapiGetUser.SetValue("XBLNR", P_DOCUMENTO);

                    BapiGetUser.Invoke(SapRfcDestination);




                    url = BapiGetUser.GetString("URL");
                    GC.Collect();
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
                GC.Collect();
            }
        }

    }
}
