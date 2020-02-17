using InternshipScannerV2.Controllers;
using InternshipScannerV2.ViewModels;
using InternshipScannerV2.Views;
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

namespace InternshipScannerV2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainView mainV= null;
        private SettingsView settingsV= null;
        public MainWindow()
        {
            InitializeComponent();
            mainV = new MainView();
            settingsV = new SettingsView();
            DataContext = mainV;
        }

        private void mainWindow_Click(object sender, RoutedEventArgs e)
        {
            DataContext = mainV;
        }

        private void settings_Click(object sender, RoutedEventArgs e)
        {
            DataContext = settingsV;
        }
    }


}
