using InternshipScannerV2.Controllers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
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

namespace InternshipScannerV2.Views
{
    /// <summary>
    /// Interaction logic for MainView.xaml
    /// </summary>
    public partial class MainView : UserControl
    {
        public bool isReadyForInput { get; set; }
        public int studentsScreened { get; set; }
        public int studentsInDK { get; set; }
        Scanner sc; 
        public MainView()
        {
            InitializeComponent();
            wbSample.Navigated += new NavigatedEventHandler(wbMain_Navigated);
            isReadyForInput = false;
            studentsScreened = 0;
            studentsInDK = 0;
            btnApprove.IsEnabled = false;
            btnDeny.IsEnabled = false;

        }
        #region WPF garbage
        void wbMain_Navigated(object sender, NavigationEventArgs e)
        {
            SetSilent(wbSample, true); // make it silent
        }

        private void txtUrl_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                wbSample.Navigate(txtUrl.Text);
        }

        private void wbSample_Navigating(object sender, System.Windows.Navigation.NavigatingCancelEventArgs e)
        {
            txtUrl.Text = e.Uri.OriginalString;
        }

        private void BrowseBack_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = ((wbSample != null) && (wbSample.CanGoBack));
        }

        private void BrowseBack_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            wbSample.GoBack();
        }

        private void BrowseForward_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = ((wbSample != null) && (wbSample.CanGoForward));
        }

        private void BrowseForward_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            wbSample.GoForward();
        }

        private void GoToPage_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void GoToPage_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            wbSample.Navigate(txtUrl.Text);
        }

        public static void SetSilent(WebBrowser browser, bool silent)
        {
            if (browser == null)
                throw new ArgumentNullException("browser");

            // get an IWebBrowser2 from the document
            IOleServiceProvider sp = browser.Document as IOleServiceProvider;
            if (sp != null)
            {
                Guid IID_IWebBrowserApp = new Guid("0002DF05-0000-0000-C000-000000000046");
                Guid IID_IWebBrowser2 = new Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E");

                object webBrowser;
                sp.QueryService(ref IID_IWebBrowserApp, ref IID_IWebBrowser2, out webBrowser);
                if (webBrowser != null)
                {
                    webBrowser.GetType().InvokeMember("Silent", BindingFlags.Instance | BindingFlags.Public | BindingFlags.PutDispProperty, null, webBrowser, new object[] { silent });
                }
            }
        }


        [ComImport, Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleServiceProvider
        {
            [PreserveSig]
            int QueryService([In] ref Guid guidService, [In] ref Guid riid, [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
        }
        #endregion


        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            string internationalStudentsExcelFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\InternationalStudentsList\International students - B&T.xlsx";
            string dataFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\Data";
            string resultsFilePath = @"D:\Google Drive\UCN Work\InternshipScannerV2\Result\";
            string resultName = "Result";
            sc = new Scanner(resultName, resultsFilePath, internationalStudentsExcelFilePath, dataFilePath, tbStatusBox);
            tbIntStudents.Content = sc.CollectInternationalEmails();
            sc.GetAllExcelFiles();
            tbIntStudents.Content = sc.ProcessExcelFiles();
            //sc.ScanStudents(tbStudentName, tbStudentEmail, tbStudentWorkPlace, tbEducation);
            sc.StepScan(tbStudentName, tbStudentEmail, tbStudentWorkPlace, tbEducation);
            updateGoogleMaps();
            isReadyForInput = true;
            btnStart.IsEnabled = false;
            ToggleRedGreenButtons();
        }

        private void Start()
        {
            
        }

        private void btnApprove_Click(object sender, RoutedEventArgs e)
        {
            if (isReadyForInput)
            {
                sc.approveStudent();
                // TODO check if true or false, to see whether or not to finish operations
                if (sc.StepScan(tbStudentName, tbStudentEmail, tbStudentWorkPlace, tbEducation))
                {
                    studentsScreened++;
                    tbStudentsScreened.Content = studentsScreened;
                    studentsInDK++;
                    tbStudentsDK.Content = studentsInDK;
                    updateGoogleMaps();
                }

                else
                {
                    ToggleRedGreenButtons();
                    sc.saveEducationResults();
                }
            }
        }

        private void btnDeny_Click(object sender, RoutedEventArgs e)
        {
            if (isReadyForInput)
            {
                sc.rejectStudent();
                //TODO Print out deny option
                if (sc.StepScan(tbStudentName, tbStudentEmail, tbStudentWorkPlace, tbEducation))
                {
                    studentsScreened++;
                    tbStudentsScreened.Content = studentsScreened;
                    updateGoogleMaps();
                }
                else
                {
                    ToggleRedGreenButtons();
                    sc.saveEducationResults();
                }

            }
        }

        private void updateGoogleMaps()
        {
            string newurl = "https://www.google.com/maps/search/" + tbStudentWorkPlace.Text;
            wbSample.Navigate(newurl);
        }

        void ToggleRedGreenButtons()
        {
            if(btnApprove.IsEnabled || btnDeny.IsEnabled)
            {
                btnApprove.IsEnabled = false;
                btnDeny.IsEnabled = false;
            }
            else
            {
                btnApprove.IsEnabled = true;
                btnDeny.IsEnabled = true;
            }
        }
    }

    
}
