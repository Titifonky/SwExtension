using System;
using LoreSoft.MathExpressions;
using LogDebugging;

namespace Outils
{
    public static class Calcul
    {
        static Calcul()
        {
            _Calcul = new MathEvaluator();
        }

        private static MathEvaluator _Calcul;

        public static VariableDictionary Variables
        {
            get { return _Calcul.Variables; }
        }

        public static Double? Evaluer(this String Expression)
        {
            try
            {
                return _Calcul.Evaluate(Expression);
            }
            catch
            { }

            return null;
        }
    }

    public static class MathX
    {
        public const Double Rad90D = 0.5 * Math.PI;
        public const Double Rad180D = Math.PI;
        public const Double Rad360D = 2 * Math.PI;

        public static Double Degree(this Double Rad)
        {
            return Rad * 180.0 / Math.PI;
        }

        public static Double Radian(this Double Deg)
        {
            return Deg * Math.PI / 180.0;
        }

        public static int Concat(this int a, int b)
        {
            return Convert.ToInt32(string.Format("{0}{1}", a, b));
        }

        public static Double eToDouble(this String s)
        {
            Double Val = 0;

            try
            {
                Val = (Double)Convert.ChangeType(s, typeof(Double));
            }
            catch
            {
                try
                {
                    Val = Double.Parse(s, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo);
                }
                catch (Exception e)
                { Log.Message(e); }

                Val = 0;
            }

            return Val;
        }

        public static int eToInteger(this String s)
        {
            int Val = 0;

            try
            {
                Val = (int)Convert.ChangeType(s, typeof(int));
            }
            catch
            {
                try
                {
                    Val = int.Parse(s, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo);
                }
                catch
                {
                    Val = 0;
                }
            }

            return Val;
        }
    }
}
