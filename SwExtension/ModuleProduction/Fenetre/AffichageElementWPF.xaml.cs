using Outils;
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
        SortedDictionary<int, Corps> ListeCorps = null;

        private CollectionView Vue = null;

        private String Campagne = "1";

        private Boolean EstInit = false;

        private void Init(SortedDictionary<int, Corps> listeCorps)
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

        public AffichageElementWPF(SortedDictionary<int, Corps> listeCorps)
        {
            Init(listeCorps);

            ListeViewTole.View = ListeViewTole.FindResource("VueDvp") as ViewBase;

            Cb_SelectCampagne.Visibility = Visibility.Collapsed;

            Vue.Filter = FiltreQteNull;
        }

        public AffichageElementWPF(SortedDictionary<int, Corps> listeCorps, int indiceCampagne)
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

            if(OnValider.IsRef())
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
                return true;

            if (c.Qte == 0)
                return false;

            return true;
        }

        private Boolean FiltreCampagne(Object objet)
        {
            Corps c = objet as Corps;
            if (c == null)
                return true;

            if (c.Campagne[Campagne.eToInteger()] == 0)
                return false;

            c.Qte = c.Campagne[Campagne.eToInteger()];
            return true;
        }

        private void Afficher_Check(object sender, RoutedEventArgs e)
        {
            if(Vue != null)
                Vue.Filter = null;
        }

        private void Masquer_Check(object sender, RoutedEventArgs e)
        {
            if (Vue != null)
                Vue.Filter = FiltreQteNull;
        }

        private void SelectCampagne_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var Campagne = Cb_SelectCampagne.SelectedValue.ToString();
            Vue.Filter = FiltreCampagne;
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
    }
}
