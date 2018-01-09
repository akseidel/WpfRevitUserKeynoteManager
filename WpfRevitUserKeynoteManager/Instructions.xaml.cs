using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfRevitUserKeynoteManager
{
    /// <summary>
    /// Interaction logic for Instructiona.xaml
    /// </summary>
    public partial class Instructions : Window
    {
        private string InfoFileName = "RevitKeyNotesExplained.rtf";

        public Instructions()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Load the help file at the very end on a thread
            // so that the app starts up right away.
            System.Threading.ThreadPool.QueueUserWorkItem(delegate { LoadUpHelpFile(); }, null);
        }

        public void DragWindow(object sender, MouseButtonEventArgs args)
        {
            // Watch out. Fatal error if not primary button!
            if (args.LeftButton == MouseButtonState.Pressed) { DragMove(); }
        }

        private void OnHelpTabSelected(object sender, RoutedEventArgs e)
        {
            if (sender is TabItem tab)
            {
                TabItem thisTab = sender as TabItem;
                switch (thisTab.Name)
                {
                    case "TabHelpClose":
                        Hide();
                        break;
                    default:
                        break;
                }
            }
        }

        private void LoadUpHelpFile()
        {
            string ExecutingAssemblyPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string fn = Helpers.CombineIntoPath(ExecutingAssemblyPath, InfoFileName);
            RTFFile rtfhelp = new RTFFile() { File = fn };
            Dispatcher.BeginInvoke((Action)(() => RichTextFileHelp.DataContext = rtfhelp));
        }

        private void Window_Activated(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke((Action)(() => TabControlMainHelp.SelectedIndex = 0));
        }

    }

    [ValueConversion(typeof(String), typeof(BitmapFrame))]
    public class PathToImage : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            Assembly ass = Assembly.GetEntryAssembly();
            string[] str = ass.GetManifestResourceNames();
            Stream stream = ass.GetManifestResourceStream(value.ToString());
            BitmapFrame bf = BitmapFrame.Create(stream);
            return bf;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }


}
