using LogDebugging;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Outils
{
    public static class MathConst
    {
        private static Double _R2 = 0;
        public static Double Racine2()
        {
            if (_R2 == 0)
                return Math.Sqrt(2);

            return _R2;
        }
    }

    public struct gPoint
    {
        public Double X;
        public Double Y;
        public Double Z;

        public gPoint(Double[] tab)
        {
            X = tab[0];
            Y = tab[1];
            Z = 0;
            if (tab.GetLength(0) > 2)
                Z = tab[2];
        }

        public gPoint(MathPoint pt)
        {
            Double[] Pt = (Double[])pt.ArrayData;
            X = Pt[0];
            Y = Pt[1];
            Z = Pt[2];
        }

        public gPoint(SketchPoint pt)
        {
            X = pt.X;
            Y = pt.Y;
            Z = pt.Z;
        }

        public gPoint(Vertex pt)
        {
            Double[] Pt = (Double[])pt.GetPoint();
            X = Pt[0];
            Y = Pt[1];
            Z = Pt[2];
        }

        public gPoint(Double x, Double y, Double z)
        {
            X = x;
            Y = y;
            Z = z;
        }

        public Boolean Comparer(gPoint pt, Double arrondi)
        {
            if (Distance(pt) < arrondi)
                return true;

            return false;
        }

        public Double Distance(gPoint pt)
        {
            return Math.Sqrt(Math.Pow(X - pt.X, 2) + Math.Pow(Y - pt.Y, 2) + Math.Pow(Z - pt.Z, 2));
        }

        public Double Distance(gSegment s)
        {
            gVecteur ba = new gVecteur(s.Start, this);
            gVecteur u = s.Vecteur;

            return ba.Vectoriel(u).Norme / u.Norme;
        }

        public gPoint Projection(gSegment s)
        {
            gVecteur ab = s.Vecteur;
            gVecteur ac = new gVecteur(s.Start, this);

            Double n = ab.Scalaire(ac) / ab.Norme;

            ab.Normaliser();
            ab.Multiplier(n);

            gPoint p = s.Start;
            p.Deplacer(ab);

            return p;
        }

        public gPoint Milieu(gPoint pt)
        {
            return new gPoint((X + pt.X) * 0.5, (Y + pt.Y) * 0.5, (Z + pt.Z) * 0.5);
        }

        public void Deplacer(gVecteur V)
        {
            X += V.X;
            Y += V.Y;
            Z += V.Z;
        }

        public gPoint Composer(gVecteur V)
        {
            return new gPoint(X + V.X, Y + V.Y, Z + V.Z);
        }

        public void Min(gPoint pt)
        {
            X = Math.Min(X, pt.X);
            Y = Math.Min(Y, pt.Y);
            Z = Math.Min(Z, pt.Z);
        }

        public void Max(gPoint pt)
        {
            X = Math.Max(X, pt.X);
            Y = Math.Max(Y, pt.Y);
            Z = Math.Max(Z, pt.Z);
        }

        public void Multiplier(Double n)
        {
            X *= n;
            Y *= n;
            Z *= n;
        }

        public void MultiplyTransfom(MathTransform trans)
        {
            MathUtility Mu = App.Sw.GetMathUtility();
            MathPoint mp = Mu.CreatePoint(new double[] { X, Y, Z });
            mp = mp.MultiplyTransform(trans);
            Double[] pt = (Double[])mp.ArrayData;
            X = pt[0];
            Y = pt[1];
            Z = pt[2];
        }

        public override string ToString()
        {
            return String.Format("X {0} Y {1} Z{2}", X, Y, Z);
        }

    }

    public struct gVecteur
    {
        public Double X;
        public Double Y;
        public Double Z;

        public gVecteur(gVecteur v)
        {
            X = v.X; Y = v.Y; Z = v.Z;
        }

        public gVecteur(Double x, Double y, Double z)
        {
            X = x; Y = y; Z = z;
        }

        public gVecteur(Double[] ar)
        {
            X = ar[0]; Y = ar[1]; Z = ar[2];
        }

        public gVecteur(gPoint a, gPoint b)
        {
            X = b.X - a.X;
            Y = b.Y - a.Y;
            Z = b.Z - a.Z;
        }

        public void Ajouter(gVecteur v)
        {
            X += v.X; Y += v.Y; Z += v.Z;
        }

        public gVecteur Compose(gVecteur v)
        {
            gVecteur C = new gVecteur(this);
            C.Ajouter(v);
            return C;
        }

        public void Multiplier(Double n)
        {
            X *= n;
            Y *= n;
            Z *= n;
        }

        public Double Scalaire(gVecteur v)
        {
            return (X * v.X) + (Y * v.Y) + (Z * v.Z);
        }

        public gVecteur Vectoriel(gVecteur v)
        {
            gVecteur prod = new gVecteur(
                Y * v.Z - Z * v.Y,
                Z * v.X - X * v.Z,
                X * v.Y - Y * v.X);

            return prod;
        }

        public Boolean EstColineaire(gVecteur v, Double arrondi, Boolean prendreEnCompteSens = true)
        {
            var result = false;

            var v1 = Unitaire();
            var v2 = v.Unitaire();

            var vn = v1.Vectoriel(v2);

            if (vn.Norme < arrondi)
            {
                if (prendreEnCompteSens)
                {
                    // Si on aditionne les deux vecteurs et que la norme est supérieur à 0
                    // c'est qu'ils ne sont pas opposé
                    if (v1.Compose(v2).Norme > arrondi)
                        result = true;
                }
                else
                    result = true;
            }

            return result;
        }

        /// <summary>
        /// Longueur du vecteur
        /// </summary>
        public Double Norme
        {
            get
            {
                return Math.Sqrt(Math.Pow(X, 2) + Math.Pow(Y, 2) + Math.Pow(Z, 2));
            }
        }

        public void Normaliser()
        {
            Double Lg = Norme;
            X /= Lg;
            Y /= Lg;
            Z /= Lg;
        }

        public gVecteur Unitaire()
        {
            Double Lg = Norme;

            gVecteur prod = new gVecteur(
                X / Lg,
                Y / Lg,
                Z / Lg);

            return prod;
        }

        public void Inverser()
        {
            Multiplier(-1);
        }

        public gVecteur Inverse()
        {
            gVecteur V = new gVecteur(X, Y, Z);
            V.Inverser();
            return V;
        }

        public Double Angle(gVecteur v)
        {
            return Math.Acos(Scalaire(v) / (Norme * v.Norme));
        }

        public Boolean RotationTrigo(gVecteur v, gVecteur normal)
        {
            gVecteur p = Vectoriel(v).Unitaire();
            p.Ajouter(normal.Unitaire());

            if (p.Norme > MathConst.Racine2())
                return true;

            return false;
        }

        public Double AngleX()
        {
            gVecteur V = new gVecteur(0, Y, Z);
            gVecteur Vo = new gVecteur(0, 0, Z);
            return Vo.Angle(V);
        }

        public Double AngleY()
        {
            gVecteur V = new gVecteur(X, 0, Z);
            gVecteur Vo = new gVecteur(0, 0, Z);
            return Vo.Angle(V);
        }

        public Double AngleZ()
        {
            gVecteur V = new gVecteur(X, Y, 0);
            gVecteur Vo = new gVecteur(0, Y, 0);
            return Vo.Angle(V);
        }

        public Double AnglePlXZ()
        {
            gVecteur V = new gVecteur(X, 0, Z);
            return V.Angle(this);
        }

        public Double AnglePlXY()
        {
            gVecteur V = new gVecteur(X, Y, 0);
            return V.Angle(this);
        }

        public Double AnglePlZY()
        {
            gVecteur V = new gVecteur(0, Y, Z);
            return V.Angle(this);
        }

        public MathVector MathVector()
        {
            MathUtility Mu = App.Sw.GetMathUtility();
            return Mu.CreateVector(new Double[] { X, Y, Z });
        }

        public override string ToString()
        {
            return String.Format("X {0} Y {1} Z{2}", X, Y, Z);
        }
    }

    public struct gSegment
    {
        public gPoint Start;
        public gPoint End;

        public gSegment(gPoint a, gPoint b)
        {
            Start = a;
            End = b;
        }

        public gSegment(gPoint a, gVecteur v)
        {
            Start = a;
            End = new gPoint
                (
                a.X + v.X,
                a.Y + v.Y,
                a.Z + v.Z
                );
        }

        public gSegment(Edge e)
        {
            Start = new gPoint((Vertex)e.GetStartVertex());
            End = new gPoint((Vertex)e.GetEndVertex());
        }

        public void Inverser()
        {
            gPoint P = Start;
            Start = End;
            End = P;
        }

        public gVecteur Inverse()
        {
            return new gVecteur(End, Start);
        }

        public Double Lg
        {
            get
            {
                return Math.Sqrt(Math.Pow(Start.X - End.X, 2) + Math.Pow(Start.Y - End.Y, 2) + Math.Pow(Start.Z - End.Z, 2));
            }
        }

        public gVecteur Vecteur
        {
            get
            {
                return new gVecteur(Start, End);
            }
        }

        public gPoint Milieu()
        {
            return new gPoint((Start.X + End.X) * 0.5, (Start.Y + End.Y) * 0.5, (Start.Z + End.Z) * 0.5);
        }

        public void OrienterDe(gSegment s)
        {
            Double d1 = Start.Distance(s);
            Double d2 = End.Distance(s);
            if (d1 > d2)
                Inverser();
        }

        public void OrienterVers(gSegment s)
        {
            Double d1 = Start.Distance(s);
            Double d2 = End.Distance(s);
            if (d1 < d2)
                Inverser();
        }

        public void MultiplyTransfom(MathTransform trans)
        {
            Start.MultiplyTransfom(trans);
            End.MultiplyTransfom(trans);
        }

        public Boolean Compare(gSegment s, Double arrondi)
        {
            if ((Start.Comparer(s.Start, arrondi) && End.Comparer(s.End, arrondi)) || (Start.Comparer(s.End, arrondi) && End.Comparer(s.Start, arrondi)))
                return true;

            return false;
        }

        public Boolean Compare(Edge e, Double arrondi)
        {
            gSegment s = new gSegment(e);

            return Compare(s, arrondi);
        }
    }

    public struct gPlan
    {
        public gPoint Origine;
        public gVecteur Normale;

        public gPlan(gPoint a, gVecteur v)
        {
            Origine = a;
            Normale = v;
            Normale.Normaliser();
        }

        public void Inverser()
        {
            Normale.Inverser();
        }

        public gPlan Inverse()
        {
            return new gPlan(Origine, Normale.Inverse());
        }

        public Boolean SurLePlan(gPoint p, Double arrondi)
        {
            var v2 = new gVecteur(Origine, p);
            v2.Normaliser();

            var val = Math.Abs(Normale.Vectoriel(v2).Norme - 1);

            // Si l'origine est sur le plan
            if (val < arrondi)
                return true;

            return false;
        }

        /// <summary>
        /// Verifie si deux plans sont identiques
        /// On peut également vérifier ou non les directions des normales
        /// </summary>
        /// <param name="p"></param>
        /// <param name="prendreEnCompteSensNormale"></param>
        /// <returns></returns>
        public Boolean SontIdentiques(gPlan p, Double arrondi, Boolean prendreEnCompteSensNormale = true)
        {
            var result = false;
            var normale = p.Normale;
            var origine = p.Origine;
            normale.Normaliser();

            // Si les normales sont colinéaires
            if (normale.EstColineaire(Normale, arrondi, prendreEnCompteSensNormale))
            {
                if (Origine.Comparer(origine, arrondi))
                    result = true;

                // On test si l'origine est sur le plan en calculant le
                // produit vectoriel de la norme avec le vecteur(Origine, origine)
                // Si la valeur est égale à 1, ces deux vecteurs sont perpendiculaire
                var v2 = new gVecteur(Origine, origine);
                v2.Normaliser();

                var val = Math.Abs(Normale.Vectoriel(v2).Norme - 1);

                if (val < arrondi)
                    result = true;
            }

            return result;
        }

        public override string ToString()
        {
            return "O " + Origine.ToString() + "\nN " + Normale.ToString();
        }
    }

    public class gPointComparer : IComparer<gPoint>
    {
        private ListSortDirection _Dir = ListSortDirection.Ascending;
        private Func<gPoint, Double> f;
        private Func<Double, Double, Boolean> t;

        private Func<Double, Double, Boolean> test()
        {
            if (_Dir == ListSortDirection.Ascending)
                return delegate (Double v1, Double v2) { return v1 > v2; };

            return delegate (Double v1, Double v2) { return v1 < v2; };
        }

        public gPointComparer() { }

        public gPointComparer(ListSortDirection dir, Func<gPoint, Double> coordonne)
        {
            _Dir = dir;

            t = test();

            f = coordonne;
        }

        public int Compare(gPoint p1, gPoint p2)
        {
            if (t(f(p1), f(p2)))
                return 1;

            if (f(p1) == f(p2))
                return 0;

            return -1;
        }

    }
}
