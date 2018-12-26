using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ModuleProduction.ModuleProduireDvp
{
    /// <summary>
    /// Logique d'interaction pour AffichageToleWPF.xaml
    /// </summary>
    public partial class AffichageToleWPF : Window
    {
        public AffichageToleWPF(SortedDictionary<int, Corps> listeCorps)
        {
            InitializeComponent();

            ListeViewTole.DataContext = listeCorps.Values;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] String NomProp = null)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(NomProp));
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
        }

        private void Annuler_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
