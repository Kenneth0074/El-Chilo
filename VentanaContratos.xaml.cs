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
// Le decimos a C# que usaremos la Libreria de Word
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Shapes;

namespace Programa_de_interconexion
{
    /// <summary>
    /// Lógica de interacción para VentanaContratos.xaml
    /// </summary>
    public partial class VentanaContratos : Window
    {
        public VentanaContratos()
        {
            InitializeComponent();
        }

        private void Boton_Regresar_Click(object sender, RoutedEventArgs e)
        {
            Programa_de_interconexion.MainWindow win2 = new Programa_de_interconexion.MainWindow();
            win2.Show();
            this.Close();
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void RadioButton_Checked_1(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Abrimos un nuevo proceso de word para abrir el documento
            var apWord = new Word.Application();
            // Le decimos a word que el tipo de archivo es .doc
            Word.Document doc;
            // Le decimos que guarde un objeto vacio por si no se encuentra el archivo
            object opc = Type.Missing;
            // Le decimos que queremos que abra word al terminar
            apWord.Visible = true;
            // Le decimos donde esta el archivo
            string ruta = @"C:\Users\ANACECILIA\Desktop\Kenneth\Pruebas de codigo botella\Botella-master\Programa de interconexion\bin\Debug\DocumentoNadie.doc";
            // Guardamos la ruta como un objeto para el metodo de la libreria
            object param = ruta;
            // Guardamos el nombre del marcador para el metodo de la libreria
            object nombre = "TxtNombreSolicitante";
            object Calle = "TxtCalleSolicitante";
            object NumeroE = "TxtNumeroExteriorSolicitante";
            object NumeroI = "TxtNumeroInteriorSolicitante";
            object CodigoPostal = "TxtCodigoPostalSolicitante";
            object Colonia = "TxtColoniaSolicitante";
            object Municipio = "TxtDelegacionMunicipioSolicitante";
            object Estado = "TxtEstadoSolicitante";
            object Telefono = "TxtTelefonoSolicitante";
            object Correo = "TxtCorreoSolicitante";
            object Fax1 = "TxtFaxSolicitante";

            object Nombre2 = "TxtNombreContacto";
            object Puesto2 = "TxtPuestoContacto";
            object Calle2 = "TxtCalleContacto";
            object NumeroE2 = "TxtNumeroExteriorContacto";
            object NumeroI2 = "TxtNumeroInteriorContacto";
            object CodigoPostal2 = "TxtCodigoPostalContacto";
            object Colonia2 = "TxtColoniaContacto";
            object Municipio2 = "TxtDelegacionMunicipioContacto";
            object Estado2 = "TxtEstadoContacto";
            object Telefono2 = "TxtTelefonoContacto";
            object Correo2 = "TxtCorreoContacto";
            object Fax2 = "TxtFaxContacto";

            object RPUTO = "TxtRPU";
            object NivelTension = "TxtNivelDeTensionSuministro";

            object Fechado = "TxtFechaEstimadaDeOperacion";
            object CapacidadBruta = "TxtCapacidadBrutaInstalada";
            object CapacidadIncrementar = "TxtCapacidadAIncrementar";
            object GeneracionPromedio = "TxtGeneracionPromedioMesEstimada";

            object Cumplimiento = "TxtManifestacionDeCumplimiento";

            object Solar = "TxtTecnologiaSolar";
            object Eolica = "TxtTecnologiaEolica";
            object Biomasa = "TxtTecnologiaBiomasa";
            object Cogenenracion = "TxtTecnologiaCogeneracion";
            object Otro = "TxtTecnologiaOtro";
            object Especificar = "TxtEspecificar";

            object CombustibleP = "TxtCombustiblePrincipal";
            object CombustibleS = "TxtCombustibleSecundario";
            object Unidades = "TxtNumeroDeUnidades";


            //Codigo Solicitud

            // Abrimos el documento pasando los datos anteriores como referencias
            doc = apWord.Documents.Open(ref param, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc, ref opc);
            //en caso que el marcador con el nombre "Texto" exista
            if (doc.Bookmarks.Exists("TxtNombreSolicitante"))
            {
                // Le decimos a word que encuentre el marcador con nombre "Texto"
                Word.Range TxtNombreSolicitante = doc.Bookmarks.get_Item(ref nombre).Range;
                // Le asignamos el valor al marcador
                TxtNombreSolicitante.Text = NombreSolicitante.Text;
                // Recreamos el marcador por si lo ocupamos de nuevo en el formato
                object nuevorango = TxtNombreSolicitante;
                doc.Bookmarks.Add("TxtNombreSolicitante", ref nuevorango);

            }
           
            if (doc.Bookmarks.Exists("TxtCalleSolicitante"))
            {

                Word.Range TxtCalleSolicitante = doc.Bookmarks.get_Item(ref Calle).Range;
                TxtCalleSolicitante.Text = CalleSolicitante.Text;
                object nuevorango = TxtCalleSolicitante;
                doc.Bookmarks.Add("TxtCalleSolicitante", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtNumeroExteriorSolicitante"))
            {
                Word.Range TxtNumeroExteriorSolicitante = doc.Bookmarks.get_Item(ref NumeroE).Range;
                TxtNumeroExteriorSolicitante.Text = NumExtSolicitante.Text;
                object nuevorango = TxtNumeroExteriorSolicitante;
                doc.Bookmarks.Add("TxtNumeroExteriorSolicitante", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtNumeroInteriorSolicitante")) //En Caso De Que No Exista
            {
                Word.Range TxtNumeroInteriorSolicitante = doc.Bookmarks.get_Item(ref NumeroI).Range;

                TxtNumeroInteriorSolicitante.Text = NumIntSolicitante.Text;

                object nuevorango = TxtNumeroInteriorSolicitante;

                doc.Bookmarks.Add("TxtNumeroInteriorSolicitante", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtCodigoPostalSolicitante"))
            {
                Word.Range TxtCodigoPostalSolicitante = doc.Bookmarks.get_Item(ref CodigoPostal).Range;
                TxtCodigoPostalSolicitante.Text = CodigoPostalSolicitante.Text;
                object nuevorango = TxtCodigoPostalSolicitante;
                doc.Bookmarks.Add("TxtCodigoPostalSolicitante", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtColoniaSolicitante"))
            {
                Word.Range TxtColoniaSolicitante = doc.Bookmarks.get_Item(ref Colonia).Range;
                TxtColoniaSolicitante.Text = ColoniaSolicitante.Text;
                object nuevorango = TxtColoniaSolicitante;
                doc.Bookmarks.Add("TxtColoniaSolicitante", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtDelegacionMunicipioSolicitante"))
            {
                Word.Range TxtDelegacionMunicipioSolicitante = doc.Bookmarks.get_Item(ref Municipio).Range;
                TxtDelegacionMunicipioSolicitante.Text = MunicipioSolicitante.Text;
                object nuevorango = TxtDelegacionMunicipioSolicitante;
                doc.Bookmarks.Add("TxtDelegacionMunicipioSolicitante", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtEstadoSolicitante"))
            {
                Word.Range TxtEstadoSolicitante = doc.Bookmarks.get_Item(ref Estado).Range;
                TxtEstadoSolicitante.Text = MunicipioSolicitante.Text;
                object nuevorango = TxtEstadoSolicitante;
                doc.Bookmarks.Add("TxtEstadoSolicitante", ref nuevorango);
            }

            if(doc.Bookmarks.Exists("TxtTelefonoSolicitante"))
            {
                Word.Range TxtTelefonoSolicitate = doc.Bookmarks.get_Item(ref Telefono).Range;
                TxtTelefonoSolicitate.Text = TelefonSolicitante.Text;
                object nuevorango = TxtTelefonoSolicitate;
                doc.Bookmarks.Add("TxtTelefonoSolicitante", ref nuevorango);
            }
       
            if (doc.Bookmarks.Exists("TxtCorreoSolicitante"))
            {
                Word.Range TxtCorreoSolicitante = doc.Bookmarks.get_Item(ref Correo).Range;
                TxtCorreoSolicitante.Text = CorreoSolicitante.Text;
                object nuevorango = TxtCorreoSolicitante;
                doc.Bookmarks.Add("TxtCorreoSolicitante", ref nuevorango);
            }
          /*  
            if (doc.Bookmarks.Exists("TxtFaxSolicitante"))
            {
                Word.Range TxtFaxSolicitante = doc.Bookmarks.get_Item(ref Fax1).Range;
                TxtFaxSolicitante.Text = Fax1.Text;
                object nuevorango = TxtFaxSolicitante;
                doc.Bookmarks.Add("TxtFaxSolicitante", ref nuevorango);
            }
            
            */
                                        //Codigo Contacto


            if (doc.Bookmarks.Exists("TxtNombreContacto"))
            {
                Word.Range TxtNombreContacto = doc.Bookmarks.get_Item(ref Nombre2).Range;
                TxtNombreContacto.Text = NombreContacto.Text;
                object nuevorango = TxtNombreContacto;
                doc.Bookmarks.Add("TxtNombreSolicitante", ref nuevorango);
            }
            
            if (doc.Bookmarks.Exists("TxtPuestoContacto"))
            {
                Word.Range TxtPuestoContacto = doc.Bookmarks.get_Item(ref Puesto2).Range;
                TxtPuestoContacto.Text = PuestoContacto.Text;
                object nuevorango = TxtPuestoContacto;
                doc.Bookmarks.Add("TxtPuestoContacto", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtCalleContacto"))
            {
                Word.Range TxtCalleContacto = doc.Bookmarks.get_Item(ref Calle2).Range;
                TxtCalleContacto.Text = CalleContacto.Text;
                object nuevorango = TxtCalleContacto;
                doc.Bookmarks.Add("TxtCalleContacto", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtNumeroExteriorContacto"))
            {
                Word.Range TxtNumeroExteriorContacto = doc.Bookmarks.get_Item(ref NumeroE2).Range;
                TxtNumeroExteriorContacto.Text = NumExtContacto.Text;
                object nuevorango = TxtNumeroExteriorContacto;
                doc.Bookmarks.Add("TxtNumeroExteriorContacto", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtNumeroInteriorContacto"))
            {
                Word.Range TxtNumeroInteriorContacto = doc.Bookmarks.get_Item(ref NumeroI2).Range;
                TxtNumeroInteriorContacto.Text = NumIntContacto.Text;
                object nuevorango = TxtNumeroInteriorContacto;
                doc.Bookmarks.Add("TxtNumeroInternoContacto", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtCodigoPostalContacto"))
            {
                Word.Range TxtCodigoPostalContacto = doc.Bookmarks.get_Item(ref CodigoPostal2).Range;
                TxtCodigoPostalContacto.Text = CodigoPostalContacto.Text;
                object nuevorango = TxtCodigoPostalContacto;
                doc.Bookmarks.Add("TxtCodigoaPostalContacto", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtDelegacionMunicipioContacto"))
            {
                Word.Range TxtDelegacionMunicipio = doc.Bookmarks.get_Item(ref Municipio2).Range;
                TxtDelegacionMunicipio.Text = MunicipioContacto.Text;
                object nuevorango = TxtDelegacionMunicipio;
                doc.Bookmarks.Add("TxtDelegacionMunicipioContacto", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtEstadoContacto"))
            {
                Word.Range TxtEstadoContacto = doc.Bookmarks.get_Item(ref Estado2).Range;
                TxtEstadoContacto.Text = EstadoContacto.Text;
                object nuevorango = TxtEstadoContacto;
                doc.Bookmarks.Add("TxtEstadoContacto", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtTelefonoContacto"))
            {
                Word.Range TxtTelefonoContacto = doc.Bookmarks.get_Item(ref Telefono2).Range;
                TxtTelefonoContacto.Text = TelefonoContacto.Text;
                object nuevorango = TxtTelefonoContacto;
                doc.Bookmarks.Add("TxtTelefonoContacto", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtCorreoContacto"))
            {
                Word.Range TxtCorreoContacto = doc.Bookmarks.get_Item(ref Correo2).Range;
                TxtCorreoContacto.Text = CorreoContacto.Text;
                object nuevorango = TxtCorreoContacto;
                doc.Bookmarks.Add("TxtCorreoContacto", ref nuevorango);

            }



            //Codigo Datos Del Serviciom

            if (doc.Bookmarks.Exists("TxtRPU"))
            {
                Word.Range TxtRPU = doc.Bookmarks.get_Item(ref RPUTO).Range;
                TxtRPU.Text = RPU.Text;
                object nuevorango = TxtRPU;
                doc.Bookmarks.Add("TxtRPU", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtNivelDeTension"))
            {
                Word.Range TxtNivelTensionSuministro = doc.Bookmarks.get_Item(ref NivelTension).Range;
                TxtNivelTensionSuministro.Text = NivelTensionSuministro.Text;
                object nuevorango = TxtNivelTensionSuministro;
                doc.Bookmarks.Add("TxtNivelDeTesion", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtFechaEstimadaDeOperacion"))
            {
                Word.Range TxtFechaEstimadaDeOperacion = doc.Bookmarks.get_Item(ref Fechado).Range;
                TxtFechaEstimadaDeOperacion.Text = FechaDeOperacion.Text;
                object nuevorango = TxtFechaEstimadaDeOperacion;
                doc.Bookmarks.Add("TxtFechaEstimadaDeOperacion", ref nuevorango);
            }

            if (doc.Bookmarks.Exists("TxtCapacidadBrutaInstalada"))
            {
                Word.Range TxtCapacidadBrutaInstalada = doc.Bookmarks.get_Item(ref CapacidadBruta).Range;
                TxtCapacidadBrutaInstalada.Text = CapacidadInstalada.Text;
                object nuevorango = TxtCapacidadBrutaInstalada;
                doc.Bookmarks.Add("TxtCapacidadBrutaInstalada", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtCapacidadAIncrementar"))
            {
                Word.Range TxtCapacidadAIncrementar = doc.Bookmarks.get_Item(ref CapacidadIncrementar).Range;
                TxtCapacidadAIncrementar.Text = CapacidadAIncrementar.Text;
                object nuevorango = TxtCapacidadAIncrementar;
                doc.Bookmarks.Add("TxtCapacidadAIncrementar", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtGeneracionPromedioMesEstimada"))
            {
                Word.Range TxtGeneracionPromedioMesEstimada = doc.Bookmarks.get_Item(ref GeneracionPromedio).Range;
                TxtGeneracionPromedioMesEstimada.Text = GeneracionPromedioMes.Text;
                object nuevorango = TxtGeneracionPromedioMesEstimada;
                doc.Bookmarks.Add("TxtGeneracionPromedioMesEstimada", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtCombustiblePrincipal"))
            {
                Word.Range TxtCombustiblePrincipal = doc.Bookmarks.get_Item(ref CombustibleP).Range;
                TxtCombustiblePrincipal.Text = CombustiblePrincipal.Text;
                object nuevorango = TxtCombustiblePrincipal;
                doc.Bookmarks.Add("TxtCombustiblePrincipal", ref nuevorango);

            }

            if (doc.Bookmarks.Exists("TxtCombustibleSecundario"))
            {
                Word.Range TxtCombustibleSecundario = doc.Bookmarks.get_Item(ref CombustibleS).Range;
                TxtCombustibleSecundario.Text = CombustibleSecundario.Text;
                object nuevorango = TxtCombustibleSecundario;
                doc.Bookmarks.Add("TxtNumeroDeUnidades", ref nuevorango);

            }
            /*
            if (doc.Bookmarks.Exists("TxtFaxContacto"))
            {
                Word.Range TxtFaxContactacto = doc.Bookmarks.get_Item(ref Fax2).Range;
                TxtFaxContactacto.Text = Fax

            if (chkConsumo.IsChecked == true)
            {
                marcadorConsumo = "X";

            }


            var apWord = Word.Application();
            Word.Document doc;
            object opc = Type.Missing;
            apWord.Visible = true;
            string ruta Application @"\Formato.dox"; */

        }

            private void NombreSolicitante_TextChanged(object sender, TextChangedEventArgs e)
            {

            }
    }
} 
