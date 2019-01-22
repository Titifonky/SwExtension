using LogDebugging;
using Macros;
using ModuleContraindreComposant;
using ModuleCreerSymetrie;
using ModuleEmpreinte;
using ModuleExportFichier.ModuleDxfDwg;
using ModuleExportFichier.ModulePdf;
using ModuleImporterInfos;
using ModuleInsererPercage;
using ModuleLierLesConfigurations;
using ModuleListerConfigComp;
using ModuleListerMateriaux;
using ModuleListerPercage;
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
using ModuleProduction.ModuleProduireBarre;
using ModuleProduction.ModuleProduireDebit;
using ModuleProduction.ModuleProduireDvp;
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
using System.Linq;
using System.Reflection;

namespace SwExtension
{
    public partial class swAddin : ISwAddin
    {
        private Boolean Reinitialiser = true;

        private ConfigModule ParametresModules = new ConfigModule("Parametres");

        private ListeMenu _eListeMenu = null;

        private List<Object> _ListePMP = new List<Object>();

        public void CreerCmdMgr()
        {
            _eListeMenu = new ListeMenu(_SwApp.GetCommandManager(_AddInCookie));
        }

        public void CreerMenusEtOnglets()
        {
            ParametresModules.AjouterParam("Reinitialiser", false, "Reinitialiser les menus & onglets");

            Reinitialiser = ParametresModules.GetParam("Reinitialiser").GetValeur<Boolean>();

            //foreach (int IdMnu in _ListeId)
            //    if (!(_CmdMgr.GetCommandGroup(IdMnu).IsRef() || Reinitialiser)) return;

            try
            {
                //==================================================================================================
                eMenu _Mnu = _eListeMenu.Add("Fonction", "Extension des fonctions Sw");

                _Mnu.AjouterCmde("Lc", typeof(PageLierLesConfigurations));
                _Mnu.AjouterCmde("Co", typeof(PageContraindreComposant));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Ip", typeof(PageInsererPercage));
                _Mnu.AjouterCmde("Ep", typeof(PageEmpreinte));
                _Mnu.AjouterCmde("Sp", typeof(PageCreerSymetrie));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Si", typeof(BoutonSelectionnerCorpsIdentiques));
                _Mnu.AjouterCmde("Lm", typeof(PageListerMateriaux));
                _Mnu.AjouterCmde("Lc", typeof(PageListerConfigComp));
                _Mnu.AjouterCmde("Lp", typeof(PageListerPercage));

                //==================================================================================================
                _Mnu = _eListeMenu.Add("Production", "Fonctions pour la production laser");

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Rd", typeof(PageRepererDossier));
                _Mnu.AjouterCmde("Pd", typeof(PageProduireDvp));
                _Mnu.AjouterCmde("Pb", typeof(PageProduireBarre));
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Cp", typeof(BoutonCommandeProfil));
                _Mnu.AjouterCmde("Ld", typeof(PageProduireDebit));
                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Gd", typeof(PageGenererConfigDvp));

                _Mnu.AjouterCmde("Ar", typeof(BoutonAfficherReperage));
                _Mnu.AjouterCmde("Cr", typeof(PageControlerRepere));
                _Mnu.AjouterCmde("Nr", typeof(BoutonNettoyerReperage));

                //==================================================================================================
                _Mnu = _eListeMenu.Add("Escalier", "Fonctions d'aide à la création d'escalier");

                _Mnu.AjouterCmde("Ie", typeof(PageInsererEsquisseConfig));
                _Mnu.AjouterCmde("Pp", typeof(PagePositionnerPlatine));
                _Mnu.AjouterCmde("Cp", typeof(PageConfigurerPlatine));
                _Mnu.AjouterCmde("Cc", typeof(PageConfigurerContreMarche));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Im", typeof(PageInsererMarches));
                _Mnu.AjouterCmde("Bm", typeof(PageBalancerMarches));

                //==================================================================================================
                _Mnu = _eListeMenu.Add("Macro", "Macro Sw");

                _Mnu.AjouterCmde("If", typeof(PageImporterInfos));
                _Mnu.AjouterCmde("OD", typeof(BoutonOuvrirDossier));
                _Mnu.AjouterCmde("AA", typeof(BoutonActiverAimantation), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde("AR", typeof(BoutonActiverRelationAuto), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Tr", typeof(BoutonToutReconstruire), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.AjouterCmde("Sc", typeof(BoutonSupprimerConfigDepliee), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde("Lm", typeof(BoutonListerMateriaux), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde("En", typeof(BoutonExclureNomenclature), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde("Ma", typeof(BoutonMAJListePiecesSoudees), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde("Ev", typeof(BoutonEnregistrerVue), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Cp", typeof(BoutonDecompterPercage));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Vr", typeof(CmdVoronoi), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("T1", typeof(Test), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);
                _Mnu.AjouterCmde("T2", typeof(Test2), swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal);

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("PM", typeof(PageParametres));
                _Mnu.AjouterCmde("Al", typeof(BoutonAfficherLogDebug));

                //==================================================================================================
                _Mnu = _eListeMenu.Add("Dessin", "Dessin Sw");
                _Mnu.AjouterCmde("Mp", typeof(BoutonMettreEnPage));
                _Mnu.AjouterCmde("Rt", typeof(BoutonRenommerToutesFeuilles));
                _Mnu.AjouterCmde("Mv", typeof(BoutonMasquerCorpsVue));
                _Mnu.AjouterCmde("Rv", typeof(BoutonRetournerDvp));
                _Mnu.AjouterCmde("Is", typeof(BoutonVueInverserStyle));
                

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Ed", typeof(PageDxfDwg));
                _Mnu.AjouterCmde("Ep", typeof(PagePdf));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Rf", typeof(BoutonRedimensionnerFeuille));
                _Mnu.AjouterCmde("Rn", typeof(BoutonRenommerFeuille));

                _Mnu.NouveauGroupe();
                _Mnu.AjouterCmde("Nt", typeof(PageInsererNote));

                ////==================================================================================================
                //_Mnu = _eListeMenu.Add("Laser", "Fonctions pour le débit laser");

                //_Mnu.NouveauGroupe();
                //_Mnu.AjouterCmde("Cd", typeof(PageCreerConfigDvp));
                //_Mnu.AjouterCmde("Dv", typeof(PageCreerDvp));

                //_Mnu.NouveauGroupe();
                //_Mnu.AjouterCmde("Eb", typeof(PageExportBarre));
                //_Mnu.AjouterCmde("Es", typeof(BoutonExportStructure));
                //_Mnu.AjouterCmde("Ld", typeof(PageListeDebit));

                //_Mnu.NouveauGroupe();
                //_Mnu.AjouterCmde("Nd", typeof(PageNumeroterDossier));
                //_Mnu.AjouterCmde("Vn", typeof(BoutonVerifierNumerotation));
                //_Mnu.AjouterCmde("Lr", typeof(PageListerRepere));

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
                ModelDoc2 Mdl = _SwApp.ActiveDoc;

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
        private CommandManager _CmdMgr = null;
        private CommandGroup _CmdGrp = null;
        private String _titre = "";
        private int _id;
        private List<List<Cmde>> ListeGrp = new List<List<Cmde>>();
        private List<Cmde> ListeCmde = null;
        private String _NomCallBackFunction = "";
        private String _NomEnableFunction = "";

        private static int _nextId = 0;
        private static int nextId { get { return ++_nextId; } }

        private int _nextIdImg = 0;
        private int nextIdImg { get { return _nextIdImg++; } }

        public int Id { get { return _id; } }

        public String Titre { get { return _titre; } }

        public eMenu(CommandManager cmdeMgr, int id, String titre, String info)
        {
            _CmdMgr = cmdeMgr;
            _id = id;
            _titre = titre;
            _NomCallBackFunction = "CallBackFunction";
            _NomEnableFunction = "EnableMethod";

            int cmdGroupErr = 0;
            _CmdGrp = cmdeMgr.CreateCommandGroup2(_id, _titre, info, info, -1, true, ref cmdGroupErr);

            NouveauGroupe();
        }

        private String NomFonction(String nomFonction, Type type)
        {
            return nomFonction + "(" + type.Name + ")";
        }

        public void AjouterCmde(String iconTexte, Type type, swCommandTabButtonTextDisplay_e positionTexte = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow)
        {
            int idCmd = _id.Concat(nextId);

            Cmde C = new Cmde(type.GetModuleTitre(), NomFonction(_NomCallBackFunction, type), NomFonction(_NomEnableFunction, type), idCmd, iconTexte, type.GetModuleTypeDocContexte(), positionTexte, nextIdImg);

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
            String CheminImage = Path.Combine(CheminDossier, _id.ToString() + _titre.Replace(' ', '_') + String.Format("Icons{0}x{0}.", HtImage) + ImageFormat.Bmp.ToString().ToLower());

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

        public void CreerMenus()
        {
            try
            {
                // Création des icons
                _CmdGrp.IconList = CreerIcons();
                _CmdGrp.MainIconList = _CmdGrp.IconList;

                Boolean sep = false;

                // Ajout des commandes du menu
                foreach (var Grp in ListeGrp)
                {
                    foreach (var Cmd in Grp)
                        Cmd.Ajouter(_CmdGrp);

                    // Ajout des séparateurs si ce n'est pas le dernier groupe
                    if (sep)
                        _CmdGrp.AddSpacer2(-1, (int)swCommandItemType_e.swMenuItem);

                    sep = true;
                }

                // On active le groupe
                _CmdGrp.HasToolbar = true;
                _CmdGrp.HasMenu = true;
                _CmdGrp.Activate();

                // Mise à jour des CommandId après activation des menus
                foreach (var listeCmd in ListeGrp)
                    foreach (var Cmd in listeCmd)
                        Cmd.SetCommandId(_CmdGrp);

                CreerOnglet();
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private void CreerOnglet()
        {
            try
            {
                foreach (eTypeDoc T in Enum.GetValues(typeof(eTypeDoc)))
                {
                    var Liste = new List<List<Cmde>>();

                    // Liste les cmds à afficher dans cet onglet pour ce type de document
                    for (int i = 0; i < ListeGrp.Count; i++)
                    {
                        var ListeCmd = new List<Cmde>();
                        var Grp = ListeGrp[i];
                        foreach (var Cmd in Grp)
                        {
                            if (Cmd.Type.HasFlag(T))
                                ListeCmd.Add(Cmd);
                        }

                        if (ListeCmd.Count > 0)
                            Liste.Add(ListeCmd);
                    }

                    // Si la liste comprend des cmds
                    if (Liste.Count > 0)
                    {
                        CommandTab cmdTab = _CmdMgr.GetCommandTab((int)Sw.eGetSwTypeDoc(T), Titre);
                        if(cmdTab.IsRef())
                            _CmdMgr.RemoveCommandTab(cmdTab);

                        cmdTab = _CmdMgr.AddCommandTab((int)Sw.eGetSwTypeDoc(T), Titre);

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

            public void Ajouter(CommandGroup CmdGrp)
            {
                int n = CmdGrp.AddCommandItem2(Titre, Position, InfoBulle, InfoBulle, IndexImage, CallbackFunction, EnableMethod, AddinId, Options);
                CommandIndex = n;
            }

            public void SetCommandId(CommandGroup CmdGrp)
            {
                CommandId = CmdGrp.CommandID[CommandIndex];
            }
        }

        private void RemoveTab()
        {
            foreach (eTypeDoc T in Enum.GetValues(typeof(eTypeDoc)))
            {
                Object[] tTab = _CmdMgr.CommandTabs((int)Sw.eGetSwTypeDoc(T));

                if (tTab.IsRef())
                {
                    foreach (CommandTab tab in tTab)
                        _CmdMgr.RemoveCommandTab(tab);
                }
            }
        }

        public void Remove()
        {
            RemoveTab();

            _CmdMgr.RemoveCommandGroup2(Id, false);

            _CmdMgr = null;
            _CmdGrp = null;
        }
    }

    public class ListeMenu : IEnumerable
    {
        private List<eMenu> _ListeMnu;

        private List<int> _ListeId;

        private CommandManager _CmdMgr;

        public ListeMenu(CommandManager cmdMgr)
        {
            _ListeMnu = new List<eMenu>();
            _ListeId = new List<int>();
            _CmdMgr = cmdMgr;
        }

        public eMenu Add(String titre, String info)
        {
            int id = _ListeMnu.Count + 1;
            eMenu mnu = new eMenu(_CmdMgr, id, titre, info);
            _ListeMnu.Add(mnu);
            _ListeId.Add(id);
            return mnu;
        }

        public eMenu this[int index] { get { return _ListeMnu[index]; } }

        public List<int> ListeId { get { return _ListeId; } }

        public int Count { get { return _ListeMnu.Count; } }

        public void CreerMenus()
        {
            foreach (eMenu mnu in _ListeMnu)
                mnu.CreerMenus();
        }

        public void RemoveMenus()
        {
            foreach (eMenu menu in _ListeMnu)
                menu.Remove();

            _CmdMgr = null;
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
