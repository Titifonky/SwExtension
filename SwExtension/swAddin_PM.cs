using LogDebugging;
using Macros;
using ModuleContraindreComposant;
using ModuleCreerSymetrie;
using ModuleEmpreinte;
using ModuleExportFichier.ModuleDxfDwg;
using ModuleExportFichier.ModulePdf;
using ModuleImporterInfos;
using ModuleInsererPercage;
using ModuleInsererPercageTole;
using ModuleLaser.ModuleCreerConfigDvp;
using ModuleLaser.ModuleCreerDvp;
using ModuleLierLesConfigurations;
using ModuleListerConfigComp;
using ModuleListerMateriaux;
using ModuleListerPercage;
using ModuleLumiere;
using ModuleMarcheConfig.ModuleConfigurerContreMarche;
using ModuleMarcheConfig.ModuleConfigurerPlatine;
using ModuleMarcheConfig.ModuleInsererEsquisseConfig;
using ModuleMarcheConfig.ModulePositionnerPlatine;
using ModuleMarchePositionner.ModuleBalancerMarches;
using ModuleMarchePositionner.ModuleInsererMarches;
using ModuleParametres;
using ModuleProduction;
using ModuleProduction.ModuleControlerRepere;
using ModuleProduction.ModuleGenererConfigDvp;
using ModuleProduction.ModuleInsererNote;
using ModuleProduction.ModuleModifierDvp;
using ModuleProduction.ModuleProduireBarre;
using ModuleProduction.ModuleProduireDebit;
using ModuleProduction.ModuleProduireDvp;
using ModuleProduction.ModuleRepereCorps;
using ModuleProduction.ModuleRepererDossier;
using ModuleVoronoi;
using Outils;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Reflection;

namespace SwExtension
{
    public partial class swAddin : ISwAddin
    {
        private Boolean Reinitialiser = true;

        private ConfigModule ParametresModules = new ConfigModule("Parametres");

        private ListeMenu _eListeMenu = null;

        private List<Object> _ListePMP = new List<Object>();

        public void CreerMenusEtOnglets(CommandManager commandManager)
        {
            _eListeMenu = new ListeMenu(commandManager);

            ParametresModules.AjouterParam("Reinitialiser", false, "Reinitialiser les menus & onglets");

            Reinitialiser = ParametresModules.GetParam("Reinitialiser").GetValeur<Boolean>();

            _eListeMenu.RemoveMenus();

            try
            {
                eMenu _Mnu;
                int d = 0;

                //==================================================================================================
                d = 10000;
                _Mnu = _eListeMenu.Add(d++, "Fonction Ext", "Extension des fonctions Sw");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Lc", typeof(PageLierLesConfigurations));
                _Mnu.AjouterCmde(d++, "Co", typeof(PageContraindreComposant));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Ip", typeof(PageInsererPercage));
                _Mnu.AjouterCmde(d++, "It", typeof(PageInsererPercageTole));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Ep", typeof(PageEmpreinte));
                _Mnu.AjouterCmde(d++, "Sp", typeof(PageCreerSymetrie));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Si", typeof(BoutonSelectionnerCorpsIdentiques));
                _Mnu.AjouterCmde(d++, "Lm", typeof(PageListerMateriaux));
                _Mnu.AjouterCmde(d++, "Lc", typeof(PageListerConfigComp));
                _Mnu.AjouterCmde(d++, "Lp", typeof(PageListerPercage));

                //==================================================================================================
                d = 20000;
                _Mnu = _eListeMenu.Add(d++, "Production", "Fonctions pour la production laser");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Rd", typeof(PageRepererDossier));
                _Mnu.AjouterCmde(d++, "Pd", typeof(PageProduireDvp));
                _Mnu.AjouterCmde(d++, "Pb", typeof(PageProduireBarre));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Cp", typeof(BoutonCommandeProfil));
                _Mnu.AjouterCmde(d++, "Ld", typeof(PageProduireDebit));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Gd", typeof(PageGenererConfigDvp));
                _Mnu.AjouterCmde(d++, "Ar", typeof(BoutonAfficherReperage));
                _Mnu.AjouterCmde(d++, "Cr", typeof(PageControlerRepere));
                _Mnu.AjouterCmde(d++, "Rc", typeof(PageRepereCorps));
                _Mnu.AjouterCmde(d++, "Nr", typeof(BoutonNettoyerReperage));
                _Mnu.AjouterCmde(d++, "Ae", typeof(BoutonAfficherMasquerEsquisseReperage));

                //==================================================================================================
                d = 30000;
                _Mnu = _eListeMenu.Add(d++, "Escalier", "Fonctions d'aide à la création d'escalier");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Ie", typeof(PageInsererEsquisseConfig));
                _Mnu.AjouterCmde(d++, "Pp", typeof(PagePositionnerPlatine));
                _Mnu.AjouterCmde(d++, "Cp", typeof(PageConfigurerPlatine));
                _Mnu.AjouterCmde(d++, "Cc", typeof(PageConfigurerContreMarche));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Im", typeof(PageInsererMarches));
                _Mnu.AjouterCmde(d++, "Bm", typeof(PageBalancerMarches));

                //==================================================================================================
                d = 40000;
                _Mnu = _eListeMenu.Add(d++, "Macro", "Macro Sw");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "If", typeof(PageImporterInfos));
                _Mnu.AjouterCmde(d++, "OD", typeof(BoutonOuvrirDossier));
                _Mnu.AjouterCmde(d++, "AA", typeof(BoutonActiverAimantation), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "AR", typeof(BoutonActiverRelationAuto), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Tr", typeof(BoutonToutReconstruire), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Sc", typeof(BoutonSupprimerConfigDepliee), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Lm", typeof(BoutonListerMateriaux), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "En", typeof(BoutonExclureNomenclature), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Ma", typeof(BoutonMAJListePiecesSoudees), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Ev", typeof(BoutonEnregistrerVue), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Ml", typeof(PageLumiere), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Nb", typeof(BoutonNettoyerBlocs), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Cp", typeof(BoutonDecompterPercage), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "Vr", typeof(CmdVoronoi), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "PM", typeof(PageParametres));
                _Mnu.AjouterCmde(d++, "Al", typeof(BoutonAfficherLogDebug));

                //==================================================================================================
                d = 50000;
                _Mnu = _eListeMenu.Add(d++, "Test", "Macro Test");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "T1", typeof(Test1), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "T2", typeof(Test2), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "T3", typeof(Test3), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "T4", typeof(Test4), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "T5", typeof(Test5), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde(d++, "T6", typeof(Test6), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "T7", typeof(Test7), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);


                //==================================================================================================
                d = 60000;
                _Mnu = _eListeMenu.Add(d++, "Dessin", "Dessin Sw");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Mp", typeof(BoutonMettreEnPage));
                _Mnu.AjouterCmde(d++, "Rt", typeof(BoutonRenommerToutesFeuilles));
                _Mnu.AjouterCmde(d++, "Mv", typeof(BoutonMasquerCorpsVue));
                _Mnu.AjouterCmde(d++, "Rv", typeof(BoutonRetournerDvp));
                _Mnu.AjouterCmde(d++, "Is", typeof(BoutonVueInverserStyle));
                _Mnu.AjouterCmde(d++, "Er", typeof(BoutonAfficherEsquisseAssemblage));
                _Mnu.AjouterCmde(d++, "Md", typeof(PageModifierDvp));
                _Mnu.AjouterCmde(d++, "Sg", typeof(BoutonSupprimerGravure));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Ed", typeof(PageDxfDwg));
                _Mnu.AjouterCmde(d++, "Ep", typeof(PagePdf));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Rf", typeof(BoutonRedimensionnerFeuille));
                _Mnu.AjouterCmde(d++, "Rn", typeof(BoutonRenommerFeuille));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Nt", typeof(PageInsererNote));
                _Mnu.AjouterCmde(d++, "Dj", typeof(BoutonDateDuJour));

                //==================================================================================================
                d = 70000;
                _Mnu = _eListeMenu.Add(d++, "Dvp", "Export des dvps");
                
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde(d++, "Cd", typeof(PageCreerConfigDvp));
                _Mnu.AjouterCmde(d++, "Dv", typeof(PageCreerDvp));

                _eListeMenu.CreerMenus();

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            ParametresModules.GetParam("Reinitialiser").SetValeur(false);
            ParametresModules.Sauver();
        }

        private Dictionary<String, Type> _DicType = new Dictionary<String, Type>();

        private Type GetTypeModule(String nomModule)
        {
            try
            {
                if (_DicType.ContainsKey(nomModule)) return _DicType[nomModule];

                foreach (Assembly a in AppDomain.CurrentDomain.GetAssemblies())
                {
                    foreach (Type t in a.GetTypes())
                    {
                        if (t.Name == nomModule)
                        {
                            _DicType.Add(nomModule, t);
                            return t;
                        }
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return null;
        }

        public void CallBackFunction(String nomModule)
        {
            try
            {
                Type t = GetTypeModule(nomModule);
                BoutonBase B = (BoutonBase)Activator.CreateInstance(t);
                ShowPage(B);
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        public int EnableMethod(String nomModule)
        {
            int arg = 1;

            try
            {
                ModelDoc2 Mdl = App.Sw.ActiveDoc;

                if (Mdl == null) return arg;

                Type Module = GetTypeModule(nomModule);
                eTypeDoc TypeDoc = Mdl.TypeDoc();


                arg = Module.GetModuleTypeDocContexte().HasFlag(TypeDoc).ToInt();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }

            return arg;
        }

        public void ShowPage<T>(T module)
            where T : BoutonBase
        {
            _ListePMP.Add(module);

            if (module.IsInit)
                module.Executer();
        }

        public void SupprimerPMP()
        {
            _ListePMP.Clear();
        }

        public void SupprimerMenus()
        {
            _eListeMenu.RemoveMenus();
        }

    }

    /// <summary>
    /// Menu SolidWorks
    /// </summary>
    public class eMenu
    {
        public CommandManager CmdMgr = null;
        public CommandGroup CmdGrp = null;
        public String Titre = "";
        public String Info = "";
        public int Id;
        private List<List<Cmde>> ListeGrp = new List<List<Cmde>>();
        private List<Cmde> ListeCmde = null;
        private String _NomCallBackFunction = "CallBackFunction";
        private String _NomEnableFunction = "EnableMethod";

        private int _nextIdImg = 0;
        private int NextIdImg { get { return _nextIdImg++; } }

        public eMenu(CommandManager cmdeMgr, int id, String titre, String info)
        {
            CmdMgr = cmdeMgr;
            Id = id;
            Titre = titre;
            Info = info;
        }

        private static String NomFonction(String nomFonction, Type type)
        {
            return nomFonction + "(" + type.Name + ")";
        }

        public void AjouterCmde(int idCmd, String iconTexte, Type type, swCommandTabButtonTextDisplay_e positionTexte = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)
        {
            Cmde C = new Cmde(type.GetModuleTitre(), NomFonction(_NomCallBackFunction, type), NomFonction(_NomEnableFunction, type), idCmd, iconTexte, type.GetModuleTypeDocContexte(), positionTexte, NextIdImg);

            ListeCmde.Add(C);
        }

        public void NouveauGroupe()
        {
            List<Cmde> L = new List<Cmde>();
            ListeGrp.Add(L);
            ListeCmde = L;
        }

        private String[] CreerIcons()
        {
            int HtImage = 40;
            String codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            String CheminDossier = Path.GetDirectoryName(Uri.UnescapeDataString(uri.Path));
            String CheminImage = Path.Combine(CheminDossier, Id.ToString() + Titre.Replace(' ', '_') + String.Format("Icons{0}x{0}.", HtImage) + ImageFormat.Bmp.ToString().ToLower());

            List<Image> ListeImg = new List<Image>();

            foreach (var G in ListeGrp)
            {
                foreach (Cmde C in G)
                {
                    ListeImg.Add(C.Icon.eConvertirEnBmp(HtImage));
                }
            }

            Image Img = new Bitmap(HtImage * ListeImg.Count, HtImage);

            using (Graphics G = Graphics.FromImage(Img))
            {
                G.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                G.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                G.SmoothingMode = SmoothingMode.AntiAlias;
                G.TextRenderingHint = TextRenderingHint.AntiAlias;

                for (int i = 0; i < ListeImg.Count; i++)
                {
                    Image I = ListeImg[i];
                    G.DrawImage(I, i * HtImage, 0);
                }
            }

            Img.Save(CheminImage, ImageFormat.Bmp);

            return new String[] { CheminImage };
        }

        public void CreerMenu()
        {
            try
            {
                int cmdGroupErr = 0;
                CmdGrp = CmdMgr.CreateCommandGroup2(Id, Titre, Info, Info, -1, true, ref cmdGroupErr);
                Log.Message("   " + Titre);
                Log.Message("      Création : " + (swCreateCommandGroupErrors)cmdGroupErr);

                // Création des icons
                CmdGrp.IconList = CreerIcons();
                CmdGrp.MainIconList = CmdGrp.IconList;

                // Nb de cmde dans le menu
                var index = 0;
                // Ajout des commandes du menu
                foreach (var listeCmd in ListeGrp)
                {
                    index += listeCmd.Count;
                    foreach (var Cmd in listeCmd)
                        Cmd.Ajouter(CmdGrp);
                }

                // On parcours les commandes à l'envers pour y inserer les séparateurs
                for (int i = ListeGrp.Count - 1; i > 0; i--)
                {
                    var listeCmd = ListeGrp[i];
                    index -= listeCmd.Count;
                    CmdGrp.AddSpacer2(index, (int)swCommandItemType_e.swMenuItem);

                    foreach (var cmd in listeCmd)
                        cmd.CommandIndex += i;
                }

                // On active le groupe
                CmdGrp.HasToolbar = true;
                CmdGrp.HasMenu = true;
                CmdGrp.Activate();

                // Mise à jour des CommandId après activation des menus
                foreach (var listeCmd in ListeGrp)
                    foreach (var Cmd in listeCmd)
                    {
                        Cmd.SetCommandId(CmdGrp);
                        Log.Message("        " + Cmd.Titre + "-> " + Cmd.CommandId + " // " + Cmd.AddinId);
                    }

                

                CmdGrp.Activate();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        public void CreerOnglet()
        {
            try
            {
                foreach (eTypeDoc T in Enum.GetValues(typeof(eTypeDoc)))
                {
                    var Liste = new List<List<Cmde>>();

                    // Liste les cmds à afficher dans cet onglet pour ce type de document
                    foreach (var listeCmd in ListeGrp)
                    {
                        var liste = new List<Cmde>();
                        foreach (var Cmd in listeCmd)
                        {
                            if (Cmd.Type.HasFlag(T))
                                liste.Add(Cmd);
                        }

                        if (liste.Count > 0)
                            Liste.Add(liste);
                    }

                    // Si la liste comprend des cmds
                    if (Liste.Count > 0)
                    {
                        CommandTab cmdTab = CmdMgr.GetCommandTab((int)Sw.eGetSwTypeDoc(T), Titre);
                        if(cmdTab.IsRef())
                            CmdMgr.RemoveCommandTab(cmdTab);

                        cmdTab = CmdMgr.AddCommandTab((int)Sw.eGetSwTypeDoc(T), Titre);

                        Boolean sep = false;

                        foreach (var ListeCmd in Liste)
                        {
                            var ListeId = new List<int>();
                            var ListePosition = new List<int>();

                            // Liste les cmds, on recupère les Id et la position des textes
                            foreach (var Cmd in ListeCmd)
                            {
                                ListeId.Add(Cmd.CommandId);
                                ListePosition.Add((int)Cmd.PositionTexte);
                            }

                            // On ajoute les cmds
                            CommandTabBox cmdBox = cmdTab.AddCommandTabBox();
                            cmdBox.AddCommands(ListeId.ToArray(), ListePosition.ToArray());

                            // Si c'est le deuxième groupe, on ajoute un séparateur
                            // !!!! On ne peut ajouter un séparateur que devant un nouveau CommandTabBox
                            if (sep)
                                cmdTab.AddSeparator(cmdBox, ListeId[0]);

                            sep = true;
                        }
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private class Cmde
        {
            public String Titre { get; private set; }
            public String Icon { get; private set; }
            public int Position { get; private set; }
            public String InfoBulle { get; private set; }
            public int IndexImage { get; private set; }
            public String CallbackFunction { get; private set; }
            public String EnableMethod { get; private set; }
            public int AddinId { get; private set; }
            public int CommandIndex { get; set; }
            public int CommandId { get; private set; }
            public int Options { get; private set; }
            public eTypeDoc Type { get; private set; }
            public swCommandTabButtonTextDisplay_e PositionTexte { get; private set; }

            public Cmde(String titre, String callbackFunction, String enableMethod, int id, String icon, eTypeDoc type, swCommandTabButtonTextDisplay_e positionTexte, int indexImage)
            {
                Titre = titre;
                Icon = icon;
                Position = -1;
                InfoBulle = titre;
                IndexImage = indexImage;
                CallbackFunction = callbackFunction;
                EnableMethod = enableMethod;
                AddinId = id;
                Options = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
                Type = type;
                PositionTexte = positionTexte;
                CommandId = -1;
            }

            public int Ajouter(CommandGroup CmdGrp)
            {
                CommandIndex = CmdGrp.AddCommandItem2(Titre, Position, InfoBulle, InfoBulle, IndexImage, CallbackFunction, EnableMethod, AddinId, Options);
                return CommandIndex;
            }

            public int SetCommandId(CommandGroup CmdGrp)
            {
                CommandId = CmdGrp.CommandID[CommandIndex];
                return CommandId;
            }
        }

        public void Remove()
        {
            CmdMgr.RemoveCommandGroup2(Id, false);

            CmdMgr = null;
            CmdGrp = null;
        }
    }

    public class ListeMenu : IEnumerable
    {
        private List<eMenu> _ListeMnu;

        public CommandManager CmdMgr;

        public ListeMenu(CommandManager cmdMgr)
        {
            _ListeMnu = new List<eMenu>();
            CmdMgr = cmdMgr;
        }

        public eMenu Add(int id, String titre, String info)
        {
            eMenu mnu = new eMenu(CmdMgr, id, titre, info);
            _ListeMnu.Add(mnu);
            return mnu;
        }

        public eMenu this[int index] { get { return _ListeMnu[index]; } }

        public int Count { get { return _ListeMnu.Count; } }

        public void CreerMenus()
        {
            _ListeMnu.Reverse();
            foreach (eMenu mnu in _ListeMnu)
                mnu.CreerMenu();

            _ListeMnu.Reverse();
            foreach (eMenu mnu in _ListeMnu)
                mnu.CreerOnglet();
        }

        public void RemoveMenus()
        {
            foreach (eMenu menu in _ListeMnu)
                menu.Remove();

            for (int i = 1; i < 80000; i++)
            {
                var cg = CmdMgr.GetCommandGroup(i);
                if (cg.IsRef())
                    Log.Message("CommandGroup " + cg.Name + "(" + i + ") supp : " + (swRemoveCommandGroupErrors)CmdMgr.RemoveCommandGroup2(i, false));
            }

            foreach (eTypeDoc T in Enum.GetValues(typeof(eTypeDoc)))
            {
                Object[] tTab = CmdMgr.CommandTabs((int)Sw.eGetSwTypeDoc(T));

                if (tTab.IsRef())
                    foreach (CommandTab tab in tTab)
                        Log.Message(tab.Name + " supp : " + CmdMgr.RemoveCommandTab(tab));
            }
        }

        public IEnumerator GetEnumerator() { return new Enumerator(this); }

        private class Enumerator : IEnumerator
        {
            private int _index = -1;
            private ListeMenu _ListMenu;

            public Enumerator(ListeMenu t) { _ListMenu = t; }

            public bool MoveNext()
            {
                if (_index < _ListMenu._ListeMnu.Count - 1)
                {
                    _index++;
                    return true;
                }
                else
                {
                    return false;
                }
            }

            public void Reset() { _index = -1; }

            public object Current { get { return _ListMenu[_index]; } }
        }
    }
}
