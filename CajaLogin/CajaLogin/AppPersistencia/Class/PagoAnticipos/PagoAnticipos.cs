using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.CierreCaja.Estructura;
using CajaIndu.AppPersistencia.Class.PagoAnticipos.Estructura;
using CajaIndu.AppPersistencia.Class.PartidasAbiertas.Estructura;

namespace CajaIndu.AppPersistencia.Class.PagoAnticipos
{
    class PagoAnticipos
    {
        public string comprobante = "";
        public string status = "";
        public string message = "";
        public string stringRfc = "";
        public int id_error = 0;
        ConexSAP connectorSap = new ConexSAP();
        public List<RETORNO> T_Retorno = new List<RETORNO>();

        public void pagoanticiposingreso(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER, string P_SERVER, string P_IDIOMA, string P_SOCIEDAD, List<DetalleViasPago> P_VIASPAGO, List<T_DOCUMENTOS> P_DOCSAPAGAR, string P_PAIS, string P_MONEDA, string P_CAJA, string P_CAJERO, string P_INGRESO, string P_APAGAR, string P_PARCIAL)
        {
            try
            {
                T_Retorno.Clear();
                comprobante = "";
                status = "";
                message = "";
                stringRfc = "";
                IRfcTable ANTICIPOS_resp;
                IRfcTable lt_PAGO_MESS;
                RETORNO retorno;
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

                    IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_PAGO_ANTICIPO");

                    //BapiGetUser.SetValue("I_BUKRS", P_SOCIEDAD);
                    BapiGetUser.SetValue("ID_CAJA", P_CAJA);
                    BapiGetUser.SetValue("PAY_CURRENCY", P_MONEDA);
                    BapiGetUser.SetValue("LAND", P_PAIS);
                    BapiGetUser.SetValue("PARCIAL", P_PARCIAL);


                    BapiGetUser.SetValue("TOTAL_FACTURAS", Convert.ToDouble(P_INGRESO));
                    BapiGetUser.SetValue("TOTAL_VIAS", Convert.ToDouble(P_APAGAR));
                    BapiGetUser.SetValue("DIFERENCIA", Convert.ToDouble(P_INGRESO) - Convert.ToDouble(P_APAGAR));

                    //BapiGetUser.SetValue("I_VBELN",P_NUMDOCSD);
                    IRfcTable GralDat = BapiGetUser.GetTable("VIAS_PAGO");
                    for (var i = 0; i < P_VIASPAGO.Count; i++)
                    {
                        GralDat.Append();
                        GralDat.SetValue("MANDT", P_VIASPAGO[i].MANDT);
                        GralDat.SetValue("LAND", P_VIASPAGO[i].LAND);
                        GralDat.SetValue("ID_COMPROBANTE", P_VIASPAGO[i].ID_COMPROBANTE);
                        GralDat.SetValue("ID_DETALLE", P_VIASPAGO[i].ID_DETALLE);
                        GralDat.SetValue("VIA_PAGO", P_VIASPAGO[i].VIA_PAGO);
                        double Monto = Convert.ToDouble(P_VIASPAGO[i].MONTO); // 100;
                        //long Monto = Convert.ToString(Convert.ToDouble(P_VIASPAGO[i].MONTO) / 100);
                        if (P_VIASPAGO[i].MONEDA == "CLP")
                        {
                            Monto = Monto / 100;
                        }
                        GralDat.SetValue("MONTO", Monto);
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
                        GralDat.SetValue("FECHA_VENC", Convert.ToDateTime(P_VIASPAGO[i].FECHA_VENC));
                        GralDat.SetValue("TEXTO_POSICION", P_VIASPAGO[i].TEXTO_POSICION);
                        GralDat.SetValue("ANEXO", P_VIASPAGO[i].ANEXO);
                        GralDat.SetValue("SUCURSAL", P_VIASPAGO[i].SUCURSAL);
                        GralDat.SetValue("NUM_CUENTA", P_VIASPAGO[i].NUM_CUENTA);
                        GralDat.SetValue("NUM_TARJETA", P_VIASPAGO[i].NUM_TARJETA);
                        GralDat.SetValue("NUM_VALE_VISTA", P_VIASPAGO[i].NUM_VALE_VISTA);
                        GralDat.SetValue("PATENTE", P_VIASPAGO[i].PATENTE);
                        GralDat.SetValue("NUM_VENTA", P_VIASPAGO[i].NUM_VENTA);
                        GralDat.SetValue("PAGARE", P_VIASPAGO[i].PAGARE);
                        GralDat.SetValue("FECHA_EMISION", Convert.ToDateTime(P_VIASPAGO[i].FECHA_EMISION));
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

                    IRfcTable GralDat2 = BapiGetUser.GetTable("DOCUMENTOS");
                    for (var i = 0; i < P_DOCSAPAGAR.Count; i++)
                    {
                        GralDat2.Append();
                        GralDat2.SetValue("MANDT", "");
                        GralDat2.SetValue("LAND", P_PAIS);
                        GralDat2.SetValue("ID_COMPROBANTE", "");
                        GralDat2.SetValue("POSICION", "");
                        GralDat2.SetValue("CLIENTE", P_DOCSAPAGAR[i].RUTCLI);
                        GralDat2.SetValue("TIPO_DOCUMENTO", P_DOCSAPAGAR[i].CLASE_DOC);
                        GralDat2.SetValue("SOCIEDAD", P_DOCSAPAGAR[i].SOCIEDAD);
                        GralDat2.SetValue("NRO_DOCUMENTO", P_DOCSAPAGAR[i].NDOCTO);
                        GralDat2.SetValue("NRO_REFERENCIA", P_DOCSAPAGAR[i].NREF);
                        GralDat2.SetValue("CAJERO_RESP", P_CAJERO);
                        GralDat2.SetValue("CAJERO_GEN", "");
                        GralDat2.SetValue("ID_CAJA", P_CAJA);
                        //GralDat2.SetValue("FECHA_COMP", "");
                        //GralDat2.SetValue("HORA", "");
                        GralDat2.SetValue("NRO_COMPENSACION", "");
                        GralDat2.SetValue("TEXTO_CABECERA", "");
                        GralDat2.SetValue("NULO", "");
                        GralDat2.SetValue("USR_ANULADOR", "");
                        GralDat2.SetValue("NRO_ANULACION", "");
                        GralDat2.SetValue("APROBADOR_ANULA", "");
                        GralDat2.SetValue("TXT_ANULACION", "");
                        GralDat2.SetValue("EXCEPCION", "");
                        GralDat2.SetValue("FECHA_DOC", Convert.ToDateTime(P_DOCSAPAGAR[i].FECHA_DOC));
                        GralDat2.SetValue("FECHA_VENC_DOC", Convert.ToDateTime(P_DOCSAPAGAR[i].FECVENCI));
                        GralDat2.SetValue("NUM_CUOTA", "");
                        double Monto = Convert.ToDouble(P_DOCSAPAGAR[i].MONTO.Trim());
                        if (P_DOCSAPAGAR[i].MONEDA == "CLP")
                        {
                            Monto = Monto / 100;
                            GralDat2.SetValue("MONTO_DOC", Convert.ToString(Monto));
                        }
                        else
                        {
                            GralDat.SetValue("MONTO_DOC", P_DOCSAPAGAR[i].MONTO);
                        }
                       // GralDat2.SetValue("MONTO_DOC", P_DOCSAPAGAR[i].MONTO);
                        GralDat2.SetValue("MONTO_DIFERENCIA", 0);
                        GralDat2.SetValue("TEXTO_EXCEPCION", "");
                        GralDat2.SetValue("PARCIAL", "");
                        // GralDat2.SetValue("TIME","");
                        GralDat2.SetValue("APROBADOR_EX", "");
                        GralDat2.SetValue("MONEDA", P_MONEDA);
                        GralDat2.SetValue("CLASE_CUENTA", "D");
                        GralDat2.SetValue("CLASE_DOC", P_DOCSAPAGAR[i].CLASE_DOC);
                        GralDat2.SetValue("NUM_CANCELACION", "");
                        GralDat2.SetValue("CME", P_DOCSAPAGAR[i].CME);
                        GralDat2.SetValue("NOTA_VENTA", "");
                        GralDat2.SetValue("CEBE", P_DOCSAPAGAR[i].CEBE);
                        GralDat2.SetValue("ACC", P_DOCSAPAGAR[i].ACC);
                    }
                    BapiGetUser.SetValue("DOCUMENTOS", GralDat2);
                   
                    BapiGetUser.Invoke(SapRfcDestination);
                   

                    ANTICIPOS_resp = BapiGetUser.GetTable("RETORNO");
                    for (int i = 0; i < ANTICIPOS_resp.Count(); i++)
                    {
                        ANTICIPOS_resp.CurrentIndex = i;
                        retorno = new RETORNO();
                        retorno.TYPE = ANTICIPOS_resp.GetString("TYPE");
                        //retorno.ID = ANTICIPOS_resp.GetString("ID");
                        retorno.CODE = ANTICIPOS_resp.GetString("CODE");
                        if (ANTICIPOS_resp.GetString("TYPE") == "S")
                        {
                            message = message + " - " + ANTICIPOS_resp.GetString("MESSAGE");
                        }
                        if (ANTICIPOS_resp.GetString("TYPE") == "E")
                        {
                            status = status + " - " + ANTICIPOS_resp.GetString("MESSAGE");
                        }
                        retorno.MESSAGE = ANTICIPOS_resp.GetString("MESSAGE");
                        retorno.LOG_NO = ANTICIPOS_resp.GetString("LOG_NO");
                        retorno.LOG_MSG_NO = ANTICIPOS_resp.GetString("LOG_MSG_NO");
                        retorno.MESSAGE_V1 = ANTICIPOS_resp.GetString("MESSAGE_V1");
                        retorno.MESSAGE_V2 = ANTICIPOS_resp.GetString("MESSAGE_V2");
                        retorno.MESSAGE_V3 = ANTICIPOS_resp.GetString("MESSAGE_V3");
                        if (ANTICIPOS_resp.GetString("MESSAGE_V4") != "")
                        {
                            comprobante = ANTICIPOS_resp.GetString("MESSAGE_V4");
                        }
                        retorno.MESSAGE_V4 = ANTICIPOS_resp.GetString("MESSAGE_V4");
                        T_Retorno.Add(retorno);
                       
                    }


                }
                GC.Collect();
            }

            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
                System.Windows.MessageBox.Show(ex.Message + ex.StackTrace);
                GC.Collect();
            }
            // return T_Retorno;
        }
    }
}
