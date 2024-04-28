using System.Data;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using System.Linq;
using ClosedXML.Excel;
using System.IO;
using System.Text.RegularExpressions;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            
            InitializeComponent();
            DataAccess.InitializeDatabase();
            DataAccess.DeleteData();
            Output.ItemsSource = DataAccess.GetData();
      
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; 
            dialog.DefaultExt = ".xlsx"; 
            dialog.Filter = "Excel documents (.xlsx)|*.xlsx"; 

            bool? result = dialog.ShowDialog();

            if (result == true)
            {

                string filename = dialog.FileName;
                txtbox1.Text = filename;
                var xls = new XLWorkbook(filename);
                var planilha = xls.Worksheets.First(w => w.Name == "dados");
                var totalLinhas = planilha.Rows().Count();

                string messageBoxText = "A plinilha possuí " + totalLinhas.ToString() + " registros que serão processados.";
                string caption = "INFORMAÇÃO";
                MessageBoxButton button = MessageBoxButton.OK;
                MessageBoxImage icon = MessageBoxImage.Information;
                MessageBoxResult messageBoxResult = MessageBox.Show(messageBoxText, caption, button, icon);

                string caminho_dir = System.Environment.CurrentDirectory;
                string path_rtf = caminho_dir + "\\rtf_files";
                string path_doc = caminho_dir + "\\doc_files";

                if (Directory.Exists(path_rtf))
                {
                    System.IO.Directory.Delete(path_rtf, true);
                    MessageBoxResult messageBoxAlertRtf1 = MessageBox.Show("Os arquivos *.rtf serão gerados em: " + path_rtf, caption, button, icon);
                    System.IO.Directory.CreateDirectory(path_rtf);
                }
                else
                {
                    System.IO.Directory.CreateDirectory(path_rtf);
                    MessageBoxResult messageBoxAlertRtf2 = MessageBox.Show("Os arquivos *.rtf serão gerados em: " + path_rtf, caption, button, icon);
                }

                if (Directory.Exists(path_doc))
                {
                    System.IO.Directory.Delete(path_doc, true);
                    MessageBoxResult messageBoxAlertDoc1 = MessageBox.Show("Os arquivos *.doc serão gerados em: " + path_doc, caption, button, icon);
                    System.IO.Directory.CreateDirectory(path_doc);
                }
                else
                {
                    System.IO.Directory.CreateDirectory(path_doc);
                    MessageBoxResult messageBoxAlertDoc2 = MessageBox.Show("Os arquivos *.doc serão gerados em: " + path_doc, caption, button, icon);
                }

                // primeira linha é o cabecalho
                for (int l = 2; l <= totalLinhas; l++)
                {
                    
                    var titulo = planilha.Cell($"A{l}").Value.ToString();
                    var rtf = planilha.Cell($"B{l}").Value.ToString();

                    string pattern = @"(?i)[^0-9a-záéíóúàèìòùâêîôûãõçàèìòù]";
                    string replacement = " ";
                    Regex rgx = new Regex(pattern);
                    string v = rgx.Replace(titulo, replacement);

                    using (TextWriter tw = new StreamWriter("rtf_files\\" + l+" - "+v+".rtf", false, Encoding.Default))
                    {
                        tw.Write(rtf);
                        tw.Close();
                    }

                    using (TextWriter tw2 = new StreamWriter("doc_files\\" + l + " - " + v + ".doc"))
                    {
                        tw2.Write(rtf);
                        tw2.Close();
                    }

                    DataAccess.AddData(titulo, path_doc + l + " - " + v + ".doc");


                }

                Output.ItemsSource = DataAccess.GetData();

            }
        }
    }
}