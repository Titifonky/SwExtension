using LogDebugging;
using System;
using Outils;

namespace SwExtension
{
    public class Info
    {
        public Boolean LogToWindowLog { get; set; }

        private DateTime _DateTimeStart = DateTime.Now;

        public Info()
        {
            LogToWindowLog = true;
        }

        protected void InitTime()
        {
            _DateTimeStart = DateTime.Now;
        }

        private static string GetSimplestTimeSpan(TimeSpan timeSpan)
        {
            var result = string.Empty;
            if (timeSpan.Days > 0)
            {
                result += string.Format(
                    @"{0:ddd\d}", timeSpan).TrimStart('0');
            }
            if (timeSpan.Hours > 0)
            {
                result += string.Format(
                    @"{0:hh\h}", timeSpan).TrimStart('0');
            }
            if (timeSpan.Minutes > 0)
            {
                result += string.Format(
                    @"{0:mm\m}", timeSpan).TrimStart('0');
            }
            if (timeSpan.Seconds >=1)
            {
                result += string.Format( @"{0:ss\s}", timeSpan).TrimStart('0');
            }
            if (timeSpan.TotalSeconds < 1)
            {
                result += "0s";
            }

            if (timeSpan.Milliseconds > 0 && (timeSpan.TotalSeconds <= 20))
            {
                result += string.Format(
                    @"{0:fff}", timeSpan).TrimStart('0');
            }
            return result;
        }

        protected void ExecuterEn(Boolean SansSautDeLigne = false)
        {
            if (!LogToWindowLog) return;

            TimeSpan t = DateTime.Now - _DateTimeStart;

            if(!SansSautDeLigne)
                WindowLog.SautDeLigne();

            WindowLog.EcrireF("Executé en {0}", GetSimplestTimeSpan(t));
        }

        protected void SautDeLigne()
        {
            if (!LogToWindowLog) return;

            WindowLog.SautDeLigne();
        }

        protected void AfficherTitre(String titreModule)
        {
            if (!LogToWindowLog) return;

            WindowLog.Ecrire("    " + titreModule.ToUpper());
            WindowLog.Ecrire("=".eRepeter(25));
            WindowLog.SautDeLigne();
        }
    }

    public abstract class Cmd : Info
    {
        protected abstract void Command();

        public void Executer()
        {
            App.MacroEnCours = true;

            InitTime();

            Command();

            ExecuterEn();

            App.MacroEnCours = false;
        }
    }
}
