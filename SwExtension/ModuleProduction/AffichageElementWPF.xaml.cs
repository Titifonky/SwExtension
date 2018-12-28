using System;
using System.Collections.Generic;
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
        public AffichageElementWPF(SortedDictionary<int, Corps> listeCorps)
        {
            InitializeComponent();

            ListeViewTole.ItemsSource = listeCorps.Values;
            ListeViewTole.View = ListeViewTole.FindResource("tileView") as ViewBase;

            CollectionView vue = (CollectionView)CollectionViewSource.GetDefaultView(ListeViewTole.ItemsSource);
            PropertyGroupDescription grpDesc = new PropertyGroupDescription("Dimension");
            vue.GroupDescriptions.Add(grpDesc);
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                this.DragMove();
            }
            catch (Exception) { }
        }

        public delegate void OnValiderEventHandler();

        public event OnValiderEventHandler OnValider;

        private void Valider_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            OnValider();
        }

        private void Annuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
