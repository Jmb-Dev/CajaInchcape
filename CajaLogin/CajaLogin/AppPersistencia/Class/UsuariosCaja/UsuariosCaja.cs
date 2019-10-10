using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware.Connector;
using CajaIndu.AppPersistencia.Class.PartidasAbiertas.Estructura;
using CajaIndu.AppPersistencia.Class.AutorizadorAnulaciones.Estructura;
using CajaIndu.AppPersistencia.Class.Connections;
using CajaIndu.AppPersistencia.Class.UsuariosCaja.Estructura;


namespace CajaIndu.AppPersistencia.Class.UsuariosCaja
{
    class UsuariosCaja
    {
        public List<USR_CAJA> ObjDatosUser = new List<USR_CAJA>();
        public List<LOG_APERTURA> LogApert = new List<LOG_APERTURA>();
        //public List<ViasPago> ViasPagoTransaccion = new List<ViasPago>();
        public List<ESTADO> Retorno = new List<ESTADO>();
        public string errormessage = "";
        string id_apertura = "";
        public string Mensaje = "";
        public string Sociedad = "";
        public string status = "";
        public string cajeroresp = "";
        ConexSAP connectorSap = new ConexSAP();


        public void usuarioscaja(string P_UNAME, string P_PASSWORD, string P_IDSISTEMA, string P_INSTANCIA, string P_MANDANTE, string P_SAPROUTER
            , string P_SERVER, string P_IDIOMA, string P_TEMPORAL, string P_MONTO, List<USR_CAJA> P_USUARIOS, string P_EQUIPO)
        {
            ObjDatosUser.Clear();
            LogApert.Clear();
            Retorno.Clear();
            errormessage = "";
            Mensaje = "";
            Sociedad = "";
            status = "";
            cajeroresp = "";

            IRfcStructure lt_USER;
            IRfcStructure lt_retorno;
            IRfcStructure ls_logapert;

            USR_CAJA USER_resp;
            ESTADO retorno_resp;
            LOG_APERTURA log_apert_resp;

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

                IRfcFunction BapiGetUser = SapRfcRepository.CreateFunction("ZSCP_RFC_ACCESO_CAJA");


                IRfcStructure GralDat = BapiGetUser.GetStructure("USR_CAJA");

                for (var i = 0; i < P_USUARIOS.Count; i++)
                {
                    //GralDat.Append();
                    GralDat.SetValue("LAND", P_USUARIOS[i].LAND);
                    GralDat.SetValue("ID_CAJA", P_USUARIOS[i].ID_CAJA);
                    GralDat.SetValue("USUARIO", P_USUARIOS[i].USUARIO);
                    GralDat.SetValue("SOCIEDAD", P_USUARIOS[i].SOCIEDAD);
                    GralDat.SetValue("NOM_CAJA", P_USUARIOS[i].NOM_CAJA);
                    GralDat.SetValue("TIPO_USUARIO", P_USUARIOS[i].TIPO_USUARIO);
                    GralDat.SetValue("USUARIO", P_USUARIOS[i].USUARIO);
                    GralDat.SetValue("WAERS", P_USUARIOS[i].WAERS);
                  

                }
                BapiGetUser.SetValue("USR_CAJA", GralDat);
                BapiGetUser.SetValue("TEMPORAL", P_TEMPORAL);
                BapiGetUser.SetValue("MONTO", P_MONTO);
                BapiGetUser.SetValue("EQUIPO", P_EQUIPO);
                // IRfcStructure GralDat = BapiGetUser.GetStructure("CONDICIONES");


                // BapiGetUser.SetValue("T_GET_DOC", GralDat);


                BapiGetUser.Invoke(SapRfcDestination);
                cajeroresp = BapiGetUser.GetString("CAJERO_RESPONSABLE");
                lt_retorno = BapiGetUser.GetStructure("ESTATUS");
                ls_logapert = BapiGetUser.GetStructure("LOG_APERTURA");

                lt_USER = BapiGetUser.GetStructure("USR_CAJA");

                
                if (lt_retorno.Count > 0)
                {
                    retorno_resp = new ESTADO();
                    for (int i = 0; i < lt_retorno.Count(); i++)
                    {
                       
                        if (i == 0)
                        {
                            status = lt_retorno.GetString("TYPE");
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
                            Mensaje = Mensaje + " - " + lt_retorno.GetString("MESSAGE")+ " " + lt_retorno.GetString("MESSAGE_V1");
                            id_apertura = lt_retorno.GetString("MESSAGE_V1");
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

                //lt_PART_ABIERTAS = BapiGetUser.GetTable("ZCLSP_TT_LISTA_DOCUMENTOS");
                if (lt_USER.Count > 0)
                {
                    //LLenamos la tabla de salida lt_DatGen
                    for (int i = 0; i < lt_USER.Count(); i++)
                    {
                        USER_resp = new USR_CAJA();
                        USER_resp.LAND = lt_USER.GetString("LAND");
                        Sociedad = lt_USER.GetString("SOCIEDAD");
                        USER_resp.SOCIEDAD = lt_USER.GetString("SOCIEDAD");
                        USER_resp.USUARIO = lt_USER.GetString("USUARIO");
                        USER_resp.ID_CAJA = lt_USER.GetString("ID_CAJA");
                        USER_resp.NOM_CAJA = lt_USER.GetString("NOM_CAJA");
                        USER_resp.TIPO_USUARIO = lt_USER.GetString("TIPO_USUARIO");
                        USER_resp.WAERS = lt_USER.GetString("WAERS");
                        ObjDatosUser.Add(USER_resp);
                        //ViasPagoTransaccion.Add(VIAS_PAGOS_resp.);

                    }
               }
                if (ls_logapert.Count > 0)
                {
                    //LLenamos la tabla de salida lt_DatGen
                    for (int i = 0; i < ls_logapert.Count(); i++)
                    {
                        log_apert_resp = new LOG_APERTURA();
                        log_apert_resp.MANDT = ls_logapert.GetString("MANDT");
                        log_apert_resp.ID_REGISTRO = ls_logapert.GetString("ID_REGISTRO");
                        log_apert_resp.LAND = ls_logapert.GetString("LAND");
                        log_apert_resp.ID_CAJA = ls_logapert.GetString("ID_CAJA");
                        log_apert_resp.USUARIO = ls_logapert.GetString("USUARIO");
                        if (ls_logapert.GetString("FECHA") != "0000-00-00")
                        {
                            log_apert_resp.FECHA = Convert.ToDateTime(ls_logapert.GetString("FECHA"));
                        }
                        if (ls_logapert.GetString("HORA") != "00:00:00")
                        {
                            log_apert_resp.HORA = Convert.ToDateTime(ls_logapert.GetString("HORA"));
                        }
                        log_apert_resp.MONTO = ls_logapert.GetString("MONTO");
                        log_apert_resp.MONEDA = ls_logapert.GetString("MONEDA");
                        log_apert_resp.TIPO_REGISTRO = ls_logapert.GetString("TIPO_REGISTRO");
                        log_apert_resp.ID_APERTURA = id_apertura;
                        log_apert_resp.TXT_CIERRE = ls_logapert.GetString("TXT_CIERRE");
                        log_apert_resp.BLOQUEO = ls_logapert.GetString("BLOQUEO");
                        log_apert_resp.USUARIO_BLOQ = ls_logapert.GetString("USUARIO_BLOQ");
                       if (i == 0)
                        LogApert.Add(log_apert_resp);
                        //ViasPagoTransaccion.Add(VIAS_PAGOS_resp.);
                       
                    }
                }
                GC.Collect();
            }
            else
            {
                errormessage = retval;
            }
        }
    }

}
