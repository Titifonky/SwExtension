using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SwExtension;
using System;

namespace ModuleProduction.ModuleModifierDvp
{
    public class CmdModifierDvp : Cmd
    {
        public ModelDoc2 MdlBase = null;
        public Boolean MettreAjourCampagne = false;
        public Boolean AfficherLignePliage = false;
        public Boolean AfficherNotePliage = false;
        public Boolean InscrireNomTole = false;
        public int TailleInscription = 5;

        protected override void Command()
        {
            try
            {
                MdlBase.AppliquerOptionsDessinLaser(AfficherNotePliage, TailleInscription);

                var Dessin = MdlBase.eDrawingDoc();

                foreach (var vue in Dessin.eFeuilleActive().eListeDesVues())
                    AppliquerOptionsVue(MdlBase, Dessin, vue);
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private void AppliquerOptionsVue(ModelDoc2 mdlBase, DrawingDoc dessin, View vue)
        {
            try
            {
                var liste = vue.ReferencedDocument.ePartDoc().eListeFonctionsDepliee();
                if (liste.Count > 0)
                    liste[0].eParcourirSousFonction(f => AppliquerOptionsPliage(mdlBase, f, vue));

                Boolean GravureExistante = false;
                if(vue.GetAnnotationCount() > 0)
                {
                    foreach (Annotation ann in vue.GetAnnotations())
                    {
                        if (ann.Layer == CONSTANTES.CALQUE_GRAVURE)
                        {
                            GravureExistante = true;
                            break;
                        }
                    }
                }

                if (!GravureExistante && InscrireNomTole && (vue.GetVisibleComponentCount() > 0))
                {
                    // Selection de la première face
                    var ListeComps = (object[])vue.GetVisibleComponents();
                    var ListeFaces = (object[])vue.GetVisibleEntities2((Component2)ListeComps[0], (int)swViewEntityType_e.swViewEntityType_Face);
                    var Face = (Entity)ListeFaces[0];
                    SelectData selData = default(SelectData);
                    selData = mdlBase.SelectionManager.CreateSelectData();
                    selData.View = vue;
                    Face.Select4(false, selData);

                    // On insère la note
                    Note Note = dessin.eModelDoc2().InsertNote(String.Format("$PRPWLD:\"{0}\"", CONSTANTES.REF_DOSSIER));
                    // On récupère le repère
                    var repere = Note.GetText();
                    // On l'insère en dur
                    Note.PropertyLinkedText = repere;

                    Note.SetTextJustification((int)swTextJustification_e.swTextJustificationCenter);

                    Annotation Annotation = Note.GetAnnotation();

                    TextFormat swTextFormat = Annotation.GetTextFormat(0);
                    swTextFormat.CharHeight = TailleInscription * 0.001;
                    Annotation.SetTextFormat(0, false, swTextFormat);

                    Annotation.Layer = "GRAVURE";
                    Annotation.SetLeader3((int)swLeaderStyle_e.swNO_LEADER, (int)swLeaderSide_e.swLS_SMART, true, true, false, false);

                    // On modifie la position du texte
                    Double[] P = (Double[])vue.Position;
                    Annotation.SetPosition2(P[0], P[1], 0);
                }
            }
            catch (Exception e)
            {
                this.LogErreur(new Object[] { e });
            }
        }

        private Boolean AppliquerOptionsPliage(ModelDoc2 mdlBase, Feature f, View vue)
        {
            if (f.Name.StartsWith(CONSTANTES.LIGNES_DE_PLIAGE))
            {
                vue.ShowSheetMetalBendNotes = AfficherNotePliage;

                f.eSelectionnerById2Dessin(mdlBase, vue);

                if (AfficherLignePliage)
                    mdlBase.UnblankSketch();
                else
                    mdlBase.BlankSketch();

                return true;
            }

            return false;
        }
    }
}


