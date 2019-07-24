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
        private List<SwCorps> ListeSwCorps = new List<SwCorps>();
        private List<Annotation> ListeAnnotations = new List<Annotation>();

        public PageRepereCorps()
        {
            try
            {
                Mt = (MathUtility)App.Sw.GetMathUtility();
                ListeCorps = MdlBase.pChargerNomenclature();

                InitSouris();

                WindowLog.EcrireF("Nb Corps : {0}", MdlBase.eComposantRacine().eListeCorps().Count);

                OnCalque += Calque;
                OnRunAfterActivation += ChargerCorps;
                OnRunAfterClose += RunAfterClose;
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private CtrlCheckBox _CheckBox_SelectionnerCorpsIdentiques;
        private CtrlCheckBox _CheckBox_SupprimerLesNotes;

        protected void Calque()
        {
            try
            {
                Groupe G;

                G = _Calque.AjouterGroupe("Options");

                _CheckBox_SelectionnerCorpsIdentiques = G.AjouterCheckBox("Selectionner les corps identiques");

                _CheckBox_SupprimerLesNotes = G.AjouterCheckBox("Supprimer les notes à la fin");
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
                Body2 CorpsBase = null;
                String MateriauCorpsBase = "";

                if (typeSel == swSelectType_e.swSelFACES)
                {
                    var f = MdlBase.eSelect_RecupererObjet<Face2>();
                    CorpsBase = (Body2)f.GetBody();
                    MateriauCorpsBase = GetMateriauCorpsBase(CorpsBase, MdlBase.eSelect_RecupererComposant());
                    typeCorps = CorpsBase.eTypeDeCorps();
                }

                if ((typeCorps == eTypeCorps.Barre) || (typeCorps == eTypeCorps.Tole))
                {
                    int Repere = 0;

                    String texteNote = String.Format("$PRPWLD:\"{0}\"", CONSTANTES.REF_DOSSIER);

                    Note Note = MdlBase.InsertNote(String.Format("$PRPWLD:\"{0}\"", CONSTANTES.REF_DOSSIER));
                    Annotation Annotation = Note.GetAnnotation();

                    Note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);
                    Annotation.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, true, true, false, false);

                    Repere = Note.GetText().Replace(CONSTANTES.PREFIXE_REF_DOSSIER, "").eToInteger();

                    if (ListeCorps.ContainsKey(Repere))
                    {
                        var corps = ListeCorps[Repere];
                        Note.PropertyLinkedText = String.Format("$PRPWLD:\"{0}\" (x{1})", CONSTANTES.REF_DOSSIER, corps.Campagne.Last().Value);
                        ListeAnnotations.Add(Annotation);

                        WindowLog.Ecrire(Repere);
                    }
                    else
                    {
                        MdlBase.eEffacerSelection();
                        Annotation.Select3(false, null);
                        MdlBase.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                    }

                    MdlBase.eEffacerSelection();

                    if (_CheckBox_SelectionnerCorpsIdentiques.IsChecked)
                    {
                        var ListeCorpsIdentiques = new List<Body2>();

                        foreach (var corps in ListeSwCorps)
                        {
                            if (MateriauCorpsBase != corps.Materiau) continue;

                            if (corps.Swcorps.eComparerGeometrie(CorpsBase) == Sw.Comparaison_e.Semblable)
                                ListeCorpsIdentiques.Add(corps.GetCorps());
                        }

                        foreach (var corps in ListeCorpsIdentiques)
                        {
                            corps.DisableHighlight = false;
                            
                        }
                    }

                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return 1;
        }

        private String GetMateriauCorpsBase(Body2 corpsBase, Component2 cpCorpsBase)
        {
            String materiauxCorpsBase = "";

            if (MdlBase.TypeDoc() == eTypeDoc.Piece)
                materiauxCorpsBase = corpsBase.eGetMateriauCorpsOuPiece(MdlBase.ePartDoc(), MdlBase.eNomConfigActive());
            else
                materiauxCorpsBase = corpsBase.eGetMateriauCorpsOuComp(cpCorpsBase);

            return materiauxCorpsBase;
        }

        private void ChargerCorps()
        {
            foreach (var comp in MdlBase.eComposantRacine().eRecListeComposant(c => { return c.TypeDoc() == eTypeDoc.Piece; }))
            {
                foreach (var Corps in comp.eListeCorps())
                    ListeSwCorps.Add(new SwCorps(comp, Corps, Corps.eGetMateriauCorpsOuComp(comp)));
            }
        }

        protected void RunAfterClose()
        {
            SupprimerEvenementsSouris();

            if (_CheckBox_SupprimerLesNotes.IsChecked)
            {
                MdlBase.eEffacerSelection();
                foreach (var Ann in ListeAnnotations)
                    Ann.Select3(true, null);

                MdlBase.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
            }
        }

        private class SwCorps
        {
            public Component2 Comp { get; private set; }
            public Body2 Swcorps { get; private set; }
            public String Materiau { get; private set; }

            public SwCorps(Component2 comp, Body2 swcorps, String materiau)
            {
                Comp = comp;
                Swcorps = swcorps;
                Materiau = materiau;
            }

            public Body2 GetCorps()
            {
                return Comp.eChercherCorps(Swcorps.Name, false);
            }
        }
    }
}
