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

        public void ReadSpss(string filename)
        {
            // Open file, can be read only and sequetial (for performance), or anything else
            using (FileStream fileStream = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10,
                                                          FileOptions.SequentialScan))
            {
                // Create the reader, this will read the file header
                SpssReader spssDataset = new SpssReader(fileStream);

                

                

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\Carlos\Desktop\output.txt"))
                {
                    // Iterate through all the varaibles
                    foreach (var variable in spssDataset.Variables)
                    {
                        // Display name and label
                        file.WriteLine("{0} - {1}", variable.Name, variable.Label);
                        // Display value-labels collection
                        foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                        {
                            file.WriteLine(" {0} - {1}", label.Key, label.Value);
                        }
                    }

                    // Iterate through all data rows in the file
                    foreach (var record in spssDataset.Records)
                    {
                        bool first = true;
                        foreach (var variable in spssDataset.Variables)
                        {
                            if (first)
                            {
                                file.Write(record.GetValue(variable));
                                first = false;
                            }
                            else
                            {
                                file.Write(string.Format(",{0}", record.GetValue(variable)));
                            }
                            //file.Write(variable.Name);
                            //file.Write(':');
                            // Use the corresponding variable object to get the values.
                            //file.Write(record.GetValue(variable));
                            // This will get the missing values as null, text with out extra spaces,
                            // and date values as DateTime.
                            // For original values, use record[variable] or record[int]
                        }
                        file.WriteLine(string.Empty);
                    }
                }
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                ReadSpss(filename);
            }
        }
    }
}
