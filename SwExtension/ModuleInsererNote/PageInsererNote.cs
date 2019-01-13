using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleInsererNote
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Inserer les notes"),
        ModuleNom("InsererNote"),
        ModuleDescription("Inserer la description des tôles et des profils.")
        ]
    public class PageInsererNote : BoutonPMPManager
    {
        private ModelDoc2 MdlBase = null;
        private DrawingDoc Dessin = null;
        private Mouse Souris = null;
        private MathUtility Mt = null;

        private Parametre AjouterMateriau;
        private Parametre ProfilCourt;
        private Parametre SautDeLigneMateriau;

        public PageInsererNote()
        {
            LogToWindowLog = false;

            try
            {
                AjouterMateriau = _Config.AjouterParam("AjouterMateriau", true, "Ajouter le matériau");
                ProfilCourt = _Config.AjouterParam("ProfilCourt", true, "Nom de profil court");
                SautDeLigneMateriau = _Config.AjouterParam("SautDeLigneMateriau", false, "Saut de ligne matériau");

                MdlBase = App.ModelDoc2;
                Dessin = MdlBase.eDrawingDoc();
                Mt = (MathUtility)App.Sw.GetMathUtility();
                InitSouris();

                OnCalque += Calque;
                OnRunAfterClose += RunAfterClose;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlCheckBox _CheckBox_AjouterMateriau;
        private CtrlCheckBox _CheckBox_ProfilCourt;
        private CtrlCheckBox _CheckBox_SautDeLigneMateriau;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options :");
                
                _CheckBox_ProfilCourt = G.AjouterCheckBox(ProfilCourt);
                _CheckBox_AjouterMateriau = G.AjouterCheckBox(AjouterMateriau);
                _CheckBox_SautDeLigneMateriau = G.AjouterCheckBox(SautDeLigneMateriau);
                _CheckBox_SautDeLigneMateriau.StdIndent();

                _CheckBox_AjouterMateriau.OnUnCheck += _CheckBox_SautDeLigneMateriau.UnCheck;
                _CheckBox_AjouterMateriau.OnIsCheck += _CheckBox_SautDeLigneMateriau.IsEnable;
                _CheckBox_AjouterMateriau.ApplyParam();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private String TextePropriete(eTypeCorps typeCorps)
        {
            String t = "";
            try
            {
                if (typeCorps == eTypeCorps.Tole)
                {
                    t += "Tôle ";

                    if (_CheckBox_AjouterMateriau.IsChecked)
                    {
                        t += "$PRPWLD:\"Matériau\"";

                        if (_CheckBox_SautDeLigneMateriau.IsChecked)
                            t += System.Environment.NewLine;
                        else
                            t += " ";
                    }

                    t += "ep$PRPWLD:\"Epaisseur de tôlerie\"";
                }
                else if (typeCorps == eTypeCorps.Barre)
                {
                    if (_CheckBox_ProfilCourt.IsChecked)
                        t += "$PRPWLD:\"ProfilCourt\"";
                    else
                        t += "$PRPWLD:\"Profil\"";

                    if (_CheckBox_AjouterMateriau.IsChecked)
                    {
                        if (_CheckBox_SautDeLigneMateriau.IsChecked)
                            t += System.Environment.NewLine;
                        else
                            t += " ";

                        t += "$PRPWLD:\"Matériau\"";
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return t;
        }

        private void InitSouris()
        {
            var vue = MdlBase.GetFirstModelView() as ModelView;

            if (vue.IsNull()) return;

            vue.EnableGraphicsUpdate = true;
            Souris = vue.GetMouse();
            AjouterEvenementsSouris();
        }

        private void AjouterEvenementsSouris()
        {
            if (Souris.IsNull()) return;

            Souris.MouseMoveNotify += Souris_MouseMoveNotify;
            Souris.MouseLBtnDownNotify += Souris_MouseLBtnDownNotify;
        }

        private void SupprimerEvenementsSouris()
        {
            if (Souris.IsNull()) return;

            Souris.MouseMoveNotify -= Souris_MouseMoveNotify;
            Souris.MouseLBtnDownNotify -= Souris_MouseLBtnDownNotify;
        }

        private Boolean BoutonDown = false;
        private Note Note = null;
        private Annotation Annotation = null;

        private int Souris_MouseLBtnDownNotify(int X, int Y, int WParam)
        {
            try
            {
                if (BoutonDown)
                {
                    Note = null;
                    Annotation = null;

                    BoutonDown = false;
                }
                else
                {
                    var typeSel = MdlBase.eSelect_RecupererTypeObjet();
                    var typeCorps = eTypeCorps.Autre;

                    if (typeSel == swSelectType_e.swSelFACES)
                    {
                        var f = MdlBase.eSelect_RecupererObjet<Face2>();
                        var b = (Body2)f.GetBody();
                        typeCorps = b.eTypeDeCorps();
                    }
                    else if (typeSel == swSelectType_e.swSelEDGES)
                    {
                        var e = MdlBase.eSelect_RecupererObjet<Edge>();
                        var b = (Body2)e.GetBody();
                        typeCorps = b.eTypeDeCorps();
                    }
                    else if (typeSel == swSelectType_e.swSelVERTICES)
                    {
                        var v = MdlBase.eSelect_RecupererObjet<Vertex>();
                        var b = (Body2)v.eListeDesFaces()[0].GetBody();
                        typeCorps = b.eTypeDeCorps();
                    }

                    if ((typeCorps == eTypeCorps.Barre) || (typeCorps == eTypeCorps.Tole))
                    {
                        Note = MdlBase.InsertNote(TextePropriete(typeCorps));
                        Annotation = Note.GetAnnotation();
                        Annotation.SetLeader3((int)swLeaderStyle_e.swBENT, (int)swLeaderSide_e.swLS_SMART, true, false, false, false);
                        BoutonDown = true;
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return 1;
        }

        private int Souris_MouseMoveNotify(int X, int Y, int WParam)
        {
            if (BoutonDown && (Note != null))
            {
                var c = CoordModele(X, Y);
                Annotation.SetPosition2(c.X, c.Y, 0);
            }

            return 1;
        }

        private ePoint CoordModele(int x, int y)
        {
            var vue = MdlBase.GetFirstModelView() as ModelView;

            var mt = vue.Transform;
            mt = mt.Inverse();

            double[] Xform = new double[16];
            object vXform = null;

            Xform[0] = 1.0;
            Xform[1] = 0.0;
            Xform[2] = 0.0;
            Xform[3] = 0.0;
            Xform[4] = 1.0;
            Xform[5] = 0.0;
            Xform[6] = 0.0;
            Xform[7] = 0.0;
            Xform[8] = 1.0;
            Xform[9] = x;
            Xform[10] = y;
            Xform[11] = 0.0;
            Xform[12] = 1.0;
            Xform[13] = 0.0;
            Xform[14] = 0.0;
            Xform[15] = 0.0;
            vXform = Xform;

            var Point = (MathTransform)Mt.CreateTransform(vXform);
            Point = Point.Multiply(mt);

            return new ePoint(Point.ArrayData[9], Point.ArrayData[10], 0);
        }

        protected void RunAfterClose()
        {
            SupprimerEvenementsSouris();
        }
    }
}
