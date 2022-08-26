using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Windows;
using ClosedXML.Excel;
using System.Diagnostics;
using System.IO;
using System.Windows.Controls;
using System.Windows.Input;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Win32;
using System.Text;

namespace GridToExcelUsingClosedXML
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            List<User> users = new List<User>();
            users.Add(new User() { Id = 1, Name = "John Doe", Birthday = new DateTime(1971, 7, 23) });
            users.Add(new User() { Id = 2, Name = "Jane Doe", Birthday = new DateTime(1974, 1, 17) });
            users.Add(new User() { Id = 3, Name = "Sammy Doe", Birthday = new DateTime(1991, 9, 2) });

            dgSimple.ItemsSource = users;
        }

        public class User
        {
            public int Id { get; set; }

            public string Name { get; set; }

            public DateTime Birthday { get; set; }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var filename = "";

            try
            {
                dgSimple.SelectionMode = DataGridSelectionMode.Extended;
                dgSimple.SelectAllCells();

                Clipboard.Clear();
                ApplicationCommands.Copy.Execute(null, dgSimple);

                var saveFileDialog = new SaveFileDialog
                {
                    FileName = filename != "" ? filename : "gpmfca-exportedDocument",
                    DefaultExt = ".csv",
                    Filter = "Common Seprated Documents (.csv)|*.csv"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var clip2 = Clipboard.GetText();
                    File.WriteAllText(saveFileDialog.FileName, clip2.Replace('\t', ','), Encoding.UTF8);
                    Process.Start(saveFileDialog.FileName);
                }

                dgSimple.UnselectAllCells();
                dgSimple.SelectionMode = DataGridSelectionMode.Single;
            }
            catch (Exception ex)
            {
            }
        }
    }
}
