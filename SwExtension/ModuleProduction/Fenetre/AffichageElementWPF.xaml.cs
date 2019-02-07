using LogDebugging;
using Outils;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace ModuleProduction
{
    /// <summary>
    /// Logique d'interaction pour AffichageToleWPF.xaml
    /// </summary>
    public partial class AffichageElementWPF : Window
    {
        ListeSortedCorps ListeCorps = null;

        private CollectionView Vue = null;

        private String Campagne = "1";

        private Boolean EstInit = false;

        private void Init(ListeSortedCorps listeCorps)
        {
            InitializeComponent();

            ListeCorps = listeCorps;

            ListeViewTole.ItemsSource = listeCorps.Values;

            Vue = (CollectionView)CollectionViewSource.GetDefaultView(ListeViewTole.ItemsSource);
            PropertyGroupDescription grpDescDimension = new PropertyGroupDescription("Dimension");
            PropertyGroupDescription grpDescMateriau = new PropertyGroupDescription("Materiau");
            Vue.GroupDescriptions.Add(grpDescMateriau);
            Vue.GroupDescriptions.Add(grpDescDimension);

            EstInit = true;
        }

        public AffichageElementWPF(ListeSortedCorps listeCorps)
        {
            Init(listeCorps);

            ListeViewTole.View = ListeViewTole.FindResource("VueDvp") as ViewBase;

            Cb_SelectCampagne.Visibility = Visibility.Collapsed;

            Vue.Filter = FiltreQteNull;
        }

        public AffichageElementWPF(ListeSortedCorps listeCorps, int indiceCampagne)
        {
            Init(listeCorps);

            ListeViewTole.View = ListeViewTole.FindResource("VueRepere") as ViewBase;

            Campagne = indiceCampagne.ToString();

            Bt_Annuler.Visibility = Visibility.Collapsed;
            Ck_Filtrer.Visibility = Visibility.Collapsed;
            Ck_Select.Visibility = Visibility.Collapsed;

            RemplirListBox();

            Vue.Filter = FiltreCampagne;
        }

        private void RemplirListBox()
        {
            var liste = ListeCorps.First().Value.Campagne.Keys;
            foreach (var index in liste)
                Cb_SelectCampagne.Items.Add(index.ToString());

            Cb_SelectCampagne.SelectedIndex = liste.Count - 1;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                this.DragMove();
            }
            catch (Exception) { }
        }

        private void Valider_Click(object sender, RoutedEventArgs e)
        {
            this.Close();

            if (OnValider.IsRef())
                OnValider();
        }

        private void Annuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public delegate void OnValiderEventHandler();

        public event OnValiderEventHandler OnValider;

        private Boolean FiltreQteNull(Object objet)
        {
            Corps c = objet as Corps;
            if (c == null)
                return false;

            if (c.Qte == 0)
                return false;

            return true;
        }

        private Boolean FiltreCampagne(Object objet)
        {
            try
            {
                Corps c = objet as Corps;
                if (c == null)
                    return false;

                if (!c.Campagne.ContainsKey(Campagne.eToInteger()))
                    return false;

                if (c.Campagne[Campagne.eToInteger()] == 0)
                    return false;
            }
            catch (Exception ex) { this.LogErreur(new Object[] { ex }); }

            return true;
        }

        private void Afficher_Check(object sender, RoutedEventArgs e)
        {
            if (Vue != null)
                Vue.Filter = null;
        }

        private void Masquer_Check(object sender, RoutedEventArgs e)
        {
            if (Vue != null)
                Vue.Filter = FiltreQteNull;
        }

        private void SelectCampagne_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (!EstInit) return;

                Campagne = Cb_SelectCampagne.SelectedValue.ToString();

                Vue.Filter = FiltreCampagne;

                foreach (Corps corps in ListeCorps.Values)
                {
                    if (!corps.Campagne.ContainsKey(Campagne.eToInteger()))
                        continue;

                    corps.Qte = corps.Campagne[Campagne.eToInteger()];
                }
            }
            catch (Exception ex) { this.LogErreur(new Object[] { ex }); }
        }

        private void Select_Check(object sender, RoutedEventArgs e)
        {
            if (!EstInit) return;

            foreach (Corps corps in ListeCorps.Values)
                corps.Dvp = true;
        }

        private void Deselect_Check(object sender, RoutedEventArgs e)
        {
            if (!EstInit) return;

            foreach (Corps corps in ListeCorps.Values)
                corps.Dvp = false;
        }

        private void SelectGroup_Check(object sender, RoutedEventArgs e)
        {
            HandleGroupCheck((CheckBox)sender, true);
        }

        private void DeselectGroup_Check(object sender, RoutedEventArgs e)
        {
            HandleGroupCheck((CheckBox)sender, false);
        }

        private void HandleGroupCheck(CheckBox sender, bool check)
        {
            var group = (CollectionViewGroup)sender.DataContext;
            HandleGroupCheckRecursive(group, check);
        }

        private void HandleGroupCheckRecursive(CollectionViewGroup group, bool check)
        {
            foreach (var itemOrGroup in group.Items)
            {
                if (itemOrGroup is CollectionViewGroup)
                {
                    var Group = itemOrGroup as CollectionViewGroup;

                    // Found a nested group - recursively run this method again
                    this.HandleGroupCheckRecursive(Group, check);
                }
                else
                {
                    var Cp = itemOrGroup as Corps;
                    if (Cp.IsRef())
                        Cp.Dvp = check;
                }
            }
        }

        private void Ouvrir_Modele_Click(object sender, RoutedEventArgs e)
        {
            var Mi = sender as MenuItem;
            var V = (Mi.Parent as ContextMenu).PlacementTarget as ListView;
            var corps = (Corps)V.SelectedItem;
            if(corps.IsRef())
                Sw.eOuvrir(corps.CheminFichierRepere);

            this.Topmost = true;
        }
    }
}
