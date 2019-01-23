using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Linq;

namespace ModuleProduction.ModuleInsererNote
{
    [ModuleTypeDocContexte(eTypeDoc.Dessin),
        ModuleTitre("Inserer les quantités"),
        ModuleNom("InsererNoteQuantite"),
        ModuleDescription("Inserer la quantité du repère.")
        ]
    public class PageInsererNote : BoutonPMPManager
    {
        private DrawingDoc Dessin = null;
        private Mouse Souris = null;
        private MathUtility Mt = null;

        private Parametre LigneAttache;
        private Parametre ModifierHtTexte;
        private Parametre HtTexte;

        private Parametre Reperage;
        private Parametre AfficherQuantite;

        private Parametre Description;
        private Parametre PrefixeTole;
        private Parametre AjouterMateriau;
        private Parametre ProfilCourt;
        private Parametre SautDeLigneMateriau;

        private ListeSortedCorps ListeCorps = new ListeSortedCorps();

        public PageInsererNote()
        {
            try
            {
                LigneAttache = _Config.AjouterParam("LigneAttache", true, "Ligne d'attache");
                ModifierHtTexte = _Config.AjouterParam("ModifierHtTexte", true, "Modifier la ht du texte");
                HtTexte = _Config.AjouterParam("HtTexte", 7, "Ht du texte en mm", "Ht du texte en mm");

                Reperage = _Config.AjouterParam("Reperage", true, "Reperage");
                AfficherQuantite = _Config.AjouterParam("AfficherQuantite", true, "Ajouter la quantité");

                Description = _Config.AjouterParam("Description", true, "Description");
                PrefixeTole = _Config.AjouterParam("PrefixeTole", true, "Prefixe tole");
                AjouterMateriau = _Config.AjouterParam("AjouterMateriau", true, "Ajouter le matériau");
                ProfilCourt = _Config.AjouterParam("ProfilCourt", true, "Nom de profil court");
                SautDeLigneMateriau = _Config.AjouterParam("SautDeLigneMateriau", false, "Saut de ligne matériau");

                Dessin = MdlBase.eDrawingDoc();
                Mt = (MathUtility)App.Sw.GetMathUtility();

                ListeCorps = MdlBase.pChargerNomenclature();
                InitSouris();

                OnCalque += Calque;
                OnRunAfterClose += RunAfterClose;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private GroupeAvecCheckBox _GroupeReperage;
        private GroupeAvecCheckBox _GroupeDescription;

        private CtrlCheckBox _CheckBox_LigneAttache;
        private CtrlCheckBox _CheckBox_ModifierHtTexte;
        private CtrlTextBox _Texte_HtTexte;

        private CtrlCheckBox _CheckBox_AfficherQuantite;

        private CtrlCheckBox _CheckBox_PrefixeTole;
        private CtrlCheckBox _CheckBox_AjouterMateriau;
        private CtrlCheckBox _CheckBox_ProfilCourt;
        private CtrlCheckBox _CheckBox_SautDeLigneMateriau;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");
                _CheckBox_LigneAttache = G.AjouterCheckBox(LigneAttache);

                _CheckBox_ModifierHtTexte = G.AjouterCheckBox(ModifierHtTexte);
                _Texte_HtTexte = G.AjouterTexteBox(HtTexte, false);
                _Texte_HtTexte.ValiderTexte += ValiderTextIsInteger;
                _Texte_HtTexte.StdIndent();

                _CheckBox_ModifierHtTexte.OnIsCheck += _Texte_HtTexte.IsEnable;
                _Texte_HtTexte.IsEnabled = _CheckBox_ModifierHtTexte.IsChecked;

                _GroupeReperage = _Calque.AjouterGroupeAvecCheckBox(Reperage);

                _CheckBox_AfficherQuantite = _GroupeReperage.AjouterCheckBox(AfficherQuantite);

                _GroupeDescription = _Calque.AjouterGroupeAvecCheckBox(Description);

                _CheckBox_PrefixeTole = _GroupeDescription.AjouterCheckBox(PrefixeTole);
                _CheckBox_ProfilCourt = _GroupeDescription.AjouterCheckBox(ProfilCourt);
                _CheckBox_AjouterMateriau = _GroupeDescription.AjouterCheckBox(AjouterMateriau);
                _CheckBox_SautDeLigneMateriau = _GroupeDescription.AjouterCheckBox(SautDeLigneMateriau);
                _CheckBox_SautDeLigneMateriau.StdIndent();

                _CheckBox_AjouterMateriau.OnUnCheck += _CheckBox_SautDeLigneMateriau.UnCheck;
                _CheckBox_AjouterMateriau.OnIsCheck += _CheckBox_SautDeLigneMateriau.IsEnable;
                _CheckBox_AjouterMateriau.ApplyParam();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
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

        private String TextePropriete(eTypeCorps typeCorps)
        {
            String t = "";
            try
            {
                if (typeCorps == eTypeCorps.Tole)
                {
                    if (_CheckBox_PrefixeTole.IsChecked)
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
                    var typeSel = MdlBase.eSelect_RecupererSwTypeObjet();
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
                        int Repere = 0;

                        String texteNote = "";
                        if(_GroupeReperage.IsChecked)
                            texteNote = String.Format("$PRPWLD:\"{0}\"", CONSTANTES.REF_DOSSIER);

                        Note = MdlBase.InsertNote(texteNote);
                        Annotation = Note.GetAnnotation();

                        if (_CheckBox_ModifierHtTexte.IsChecked)
                        {
                            TextFormat swTextFormat = Annotation.GetTextFormat(0);
                            swTextFormat.CharHeight = Math.Max(1, _Texte_HtTexte.Text.eToInteger()) * 0.001;
                            Annotation.SetTextFormat(0, false, swTextFormat);
                        }

                        if (!_CheckBox_LigneAttache.IsChecked)
                        {
                            Note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);
                            Annotation.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, true, true, false, false);
                        }

                        if (_GroupeReperage.IsChecked)
                        {
                            if (_CheckBox_AfficherQuantite.IsChecked)
                            {
                                Repere = Note.GetText().Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "").eToInteger();

                                if (ListeCorps.ContainsKey(Repere))
                                {
                                    var corps = ListeCorps[Repere];
                                    Note.PropertyLinkedText = String.Format("$PRPWLD:\"{0}\" (x{1})", CONSTANTES.REF_DOSSIER, corps.Campagne.Last().Value);
                                }
                            }
                        }

                        if (_GroupeDescription.IsChecked)
                        {
                            String texte = Note.PropertyLinkedText;

                            if (!String.IsNullOrWhiteSpace(texte))
                                texte += System.Environment.NewLine;

                            texte += TextePropriete(typeCorps);

                            Note.PropertyLinkedText = texte;
                        }

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
