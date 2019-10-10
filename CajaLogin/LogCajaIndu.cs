using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Data;

namespace CajaIndigo
{
    class LogCajaIndigo
    {
        public static void EscribeLogCaja(DateTime TimeNow, string User, string IdCaja, string NomCaja, string Acción)
        {
            string DirectorioLog = System.IO.Path.GetTempPath();
            DirectorioLog = DirectorioLog + "inchcapeLog\\"; //
            if (System.IO.Directory.Exists(DirectorioLog) == false)
            {
                System.IO.Directory.CreateDirectory(DirectorioLog);
            }
            DirectorioLog = DirectorioLog + "LogCajaInchcape.txt";

            using (StreamWriter writer = new StreamWriter(DirectorioLog, true))

            {
                int size = 0;
                string TimeActual = Convert.ToString(TimeNow);
                size = TimeActual.Length;

                do
                {
                    size = TimeActual.Length;
                    TimeActual = TimeActual + " ";
                } while (size < 26);
                size = 0;
                do
                {
                    size = IdCaja.Length;
                    IdCaja = IdCaja + " ";
                } while (size < 8);
                size = 0;
                do
                {
                    size = NomCaja.Length;
                    NomCaja = NomCaja + " ";
                } while (size < 30);
                size = 0;
                do
                {
                    size = User.Length;
                    User = User + " ";
                } while (size < 26);

                writer.WriteLine();
                writer.Write(TimeActual);
                writer.Write(IdCaja);
                writer.Write(NomCaja);
                writer.Write(User);
                writer.Write(Acción);

            }
        }
    }
}


