using System;
using System.Collections.Generic;
using System.Linq;
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

namespace Programa_de_interconexion
{
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Solicitud_Click(object sender, RoutedEventArgs e)
        {
            VentanaSolicitud win2 = new VentanaSolicitud();
            win2.Show();
            this.Close();
        }

        private void Fichas_Click(object sender, RoutedEventArgs e)
        {
            VentanaFichasTecnicas win2 = new VentanaFichasTecnicas();
            win2.Show();
            this.Close();
        }

        private void Croquis_Click(object sender, RoutedEventArgs e)
        {
            VentanaCroquis win2 = new VentanaCroquis();
            win2.Show();
            this.Close();
        }

        private void Diagrama_Click(object sender, RoutedEventArgs e)
        {
            VentanaDiagrama win2 = new VentanaDiagrama();
            win2.Show();
            this.Close();
        }

        private void Constancia_Click(object sender, RoutedEventArgs e)
        {
            VentanaConstancia win2 = new VentanaConstancia();
            win2.Show();
            this.Close();
        }

        private void Fotos_Click(object sender, RoutedEventArgs e)
        {
            VentanaFotos win2 = new VentanaFotos();
             win2.Show();
             this.Close();
        }

        private void RFC_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.sat.gob.mx/tramites/28753/obten-tu-rfc-con-la-clave-unica-de-registro-de-poblacion-(curp)");
        }

        private void Identificacion_Click(object sender, RoutedEventArgs e)
        {
            VentanaIdentificacion win2 = new VentanaIdentificacion();
            win2.Show();
            this.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            VentanaContratos win2 = new VentanaContratos();
            win2.Show();
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            VentanaContratos win2 = new VentanaContratos();
            win2.Show();
            this.Close();
        }
    }
}
