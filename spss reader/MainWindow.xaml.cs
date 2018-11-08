using Microsoft.Win32;
using SpssLib.DataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

namespace spss_reader
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

        public void ReadSpss(string filename, char separator)
        {
            // Open file, can be read only and sequetial (for performance), or anything else
            using (FileStream fileStream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10,
                                                          FileOptions.SequentialScan))
            {
                // Create the reader, this will read the file header
                SpssReader spssDataset = new SpssReader(fileStream);


                SaveFileDialog dlg = new SaveFileDialog
                {
                    FileName = filename.Replace(".sav", string.Empty),
                    DefaultExt = ".csv",
                    Filter = "csv documents (.csv)|*.csv"
                };

                dlg.ShowDialog();

                if (dlg.FileName != "")
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(dlg.FileName))
                    {
                        bool first = true;
                        foreach (var variable in spssDataset.Variables)
                        {
                            if (first)
                            {
                                file.Write("{0}", variable.Name);
                                first = false;
                            }
                            else
                            {
                                file.Write("{0}{1}", separator, variable.Name);
                            }
                        }
                        file.WriteLine(string.Empty);

                        // Iterate through all data rows in the file
                        foreach (var record in spssDataset.Records)
                        {
                            first = true;
                            foreach (var variable in spssDataset.Variables)
                            {
                                if (first)
                                {
                                    file.Write(record.GetValue(variable));
                                    first = false;
                                }
                                else
                                {
                                    file.Write(string.Format("{0}{1}", separator, record.GetValue(variable)));
                                }
                            }
                            file.WriteLine(string.Empty);
                        }
                    }
                }
                else
                {
                    return;
                }

                if (GenearaMetadatos.IsChecked == true)
                {
                    SaveFileDialog dlg2 = new SaveFileDialog
                    {
                        DefaultExt = ".txt",
                        Filter = "text documents (.txt)|*.txt",
                        FileName = dlg.FileName.Replace(".csv", "_metadata.txt")
                    };

                    dlg2.ShowDialog();

                    if (dlg2.FileName != "")
                    {
                        using (System.IO.StreamWriter file = new System.IO.StreamWriter(dlg2.FileName))
                        {
                            // Iterate through all the varaibles
                            foreach (var variable in spssDataset.Variables)
                            {
                                // Display name and label
                                file.WriteLine("{0} - {1}", variable.Name, variable.Label);
                                // Display value-labels collection
                                foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                                {
                                    file.WriteLine("{0} - {1}", label.Key, label.Value);
                                }
                                file.WriteLine(string.Empty);
                                file.WriteLine(string.Empty);
                            }
                        }
                    }
                }

                MessageBox.Show("Finalizado exitosamente", "Finalizado exitosamente", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ButtonAbrirSav_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            bool? result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                string filename = dlg.FileName;
                ReadSpss(filename, GetSeparator());
            }
        }

        private void ButtonAbrirMultiplesSav_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Spss Files (SAV)|*.SAV"
            };
            bool? result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                string[] filenames = dlg.FileNames;
                foreach(string filename in filenames)
                {
                    ReadSpss(filename, GetSeparator());
                }
            }
        }

        private char GetSeparator()
        {
            if (SeparatorOptions.SelectedIndex == 0)
            {
                return ',';
            }
            else if (SeparatorOptions.SelectedIndex == 0)
            {
                return ';';
            }
            else
            {
                return '|';
            }
        }
    }
}
