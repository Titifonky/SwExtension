using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleProduction.ModuleRepereCorps
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece),
        ModuleTitre("Repère du corps"),
        ModuleNom("RepereCorps"),
        ModuleDescription("Retrouve le repère du corps.")
        ]
    public class PageRepereCorps : BoutonPMPManager
    {
        private Mouse Souris = null;
        private MathUtility Mt = null;
        private ListeSortedCorps ListeCorps = new ListeSortedCorps();
        private List<Annotation> ListeAnnotations = new List<Annotation>();

        public PageRepereCorps()
        {
            try
            {
                Mt = (MathUtility)App.Sw.GetMathUtility();
                ListeCorps = MdlBase.pChargerNomenclature();

                InitSouris();

                OnCalque += Calque;
                OnRunAfterClose += RunAfterClose;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlCheckBox _CheckBox_SupprimerLesNotes;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_SupprimerLesNotes = G.AjouterCheckBox("Supprimer à la fin");
                _CheckBox_SupprimerLesNotes.IsChecked = true;
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

            Souris.MouseLBtnDownNotify += Souris_MouseLBtnDownNotify;
        }

        private void SupprimerEvenementsSouris()
        {
            if (Souris.IsNull()) return;

            Souris.MouseLBtnDownNotify -= Souris_MouseLBtnDownNotify;
        }

        private int Souris_MouseLBtnDownNotify(int X, int Y, int WParam)
        {
            try
            {
                var typeSel = MdlBase.eSelect_RecupererSwTypeObjet();
                var typeCorps = eTypeCorps.Autre;

                if (typeSel == swSelectType_e.swSelFACES)
                {
                    var f = MdlBase.eSelect_RecupererObjet<Face2>();
                    var b = (Body2)f.GetBody();
                    typeCorps = b.eTypeDeCorps();
                }

                if ((typeCorps == eTypeCorps.Barre) || (typeCorps == eTypeCorps.Tole))
                {
                    int Repere = 0;

                    String texteNote = String.Format("$PRPWLD:\"{0}\"", CONSTANTES.REF_DOSSIER);

                    Note Note = MdlBase.InsertNote(String.Format("$PRPWLD:\"{0}\"", CONSTANTES.REF_DOSSIER));
                    Annotation Annotation = Note.GetAnnotation();
                    ListeAnnotations.Add(Annotation);

                    Note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);
                    Annotation.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, true, true, false, false);

                    Repere = Note.GetText().Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "").eToInteger();

                    if (ListeCorps.ContainsKey(Repere))
                    {
                        var corps = ListeCorps[Repere];
                        Note.PropertyLinkedText = String.Format("$PRPWLD:\"{0}\" (x{1})", CONSTANTES.REF_DOSSIER, corps.Campagne.Last().Value);
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

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

            MdlBase.eEffacerSelection();
            foreach (var Ann in ListeAnnotations)
                Ann.Select3(true, null);

            MdlBase.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
        }
    }
}
