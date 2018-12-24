using LogDebugging;
using Outils;
using SwExtension;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ModuleParametres
{
    [ModuleTypeDocContexte(eTypeDoc.Assemblage | eTypeDoc.Piece | eTypeDoc.Dessin),
        ModuleTitre("Parametres des modules"),
        ModuleNom("ParametresModule"),
        ModuleDescription("Parametrer les modules")
        ]
    public class PageParametres : BoutonPMPManager
    {
        public PageParametres()
        {
            OnCalque += Calque;
            OnRunAfterActivation += delegate { _TextComboBox.SelectedIndex = 0; };
            OnRunAfterClose += RunAfterClose;
        }

        private Dictionary<String, ConfigModule> _ListeConfigModule = new Dictionary<String, ConfigModule>();

        private Groupe _GroupeModule;
        private CtrlTextComboBox _TextComboBox;
        private List<CtrlCheckBox> _ListeCheckBox = new List<CtrlCheckBox>();
        private List<CtrlTextBox> _ListeTextBox = new List<CtrlTextBox>();

        private static readonly int _MaxCtrl = 30;

        protected void Calque()
        {
            try
            {
                var G = _Calque.AjouterGroupe("Selectionner le module");

                _TextComboBox = G.AjouterTextComboBox("");
                _TextComboBox.OnSelectionChanged += SelectionChanged;

                _GroupeModule = _Calque.AjouterGroupe("Module");

                for (int i = 0; i < _MaxCtrl; i++)
                {
                    var cb = _GroupeModule.AjouterCheckBox("Titre");
                    cb.Visible = false;
                    _ListeCheckBox.Add(cb);

                    var tb = _GroupeModule.AjouterTexteBox("Titre", "Intitule");
                    tb.Visible = false;
                    _ListeTextBox.Add(tb);
                }

                foreach (var nom in ConfigModule.ListeNomConfigModule)
                {
                    ConfigModule M = new ConfigModule(nom);
                    M.ChargerParametreBrut();

                    if (M.ListeParametre.Count > 0)
                        _ListeConfigModule.Add(nom, M);
                }

                _TextComboBox.Liste = _ListeConfigModule.Keys.ToList();

            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        private void SelectionChanged(Object sender, int Item)
        {
            try
            {
                ConfigModule M = _ListeConfigModule[((CtrlTextComboBox)sender).Text];
                _GroupeModule.Titre = M.Intitule;

                var ListeParams = M.ListeParametre;

                for (int i = 0; i < _MaxCtrl; i++)
                {
                    if (i < ListeParams.Count)
                    {
                        var P = ListeParams[i];

                        CtrlCheckBox c = _ListeCheckBox[i];
                        CtrlTextBox t = _ListeTextBox[i];
                        if (P.Type == typeof(Boolean))
                        {
                            c.Visible = true;
                            t.Visible = false;
                            c.Param = P;
                            c.Caption = P.Intitule;
                            c.ApplyParam();
                        }
                        else if (P.Type == typeof(String))
                        {
                            c.Visible = false;
                            t.Visible = true;
                            t.Param = P;
                            t.LabelText = P.Intitule;
                            t.ApplyParam();
                        }
                    }
                    else
                    {
                        _ListeCheckBox[i].Visible = false;
                        _ListeTextBox[i].Visible = false;
                    }
                }
            }
            catch (Exception e)
            { this.LogMethode(new Object[] { e }); }
        }

        protected void RunAfterClose()
        {
            foreach(ConfigModule M in _ListeConfigModule.Values)
                M.Sauver();
        }

    }
}
