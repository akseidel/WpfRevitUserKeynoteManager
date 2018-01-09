using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace WpfRevitUserKeynoteManager
{
    /// <summary>
    /// Interaction logic for FormMsg.xaml 
    /// </summary>
    public partial class FormMsgWPF : Window
    {
        Brush ClrA = ColorExt.ToBrush(System.Drawing.Color.AliceBlue);
        Brush ClrB = ColorExt.ToBrush(System.Drawing.Color.Cornsilk);
        string _purpose;
        string _apath;
        bool _useTimer;
        MessageBoxResult _theResult;
     
        public string ThePath {
            get { return _apath; }
        }

        public MessageBoxResult TheResult {
            get { return _theResult; }
        }

        /// <summary>
        /// A messagebox replacement.
        /// </summary>
        /// <param name="APath"></param>
        /// <param name="reasonMode"></param>
        /// <param name="useTimer"></param>
        public FormMsgWPF(string APath, int reasonMode, bool useTimer = false)
        {
            InitializeComponent();
            Top = Properties.Settings.Default.Manager_Top + 100;
            Left = Properties.Settings.Default.Manager_Left + Properties.Settings.Default.Manager_Width/2 - this.Width;
            
            _apath = APath;
            _useTimer = useTimer;
            switch (reasonMode)
            {
                case 0:  // ask for path
                    UsersInput.Text = _apath;
                    butNo.Visibility = Visibility.Collapsed;
                    break;
                case 1: // ask for filename
                    butNo.Visibility = Visibility.Collapsed;
                    break;
                case 2: // ask for yes/no
                    UsersInput.Visibility = Visibility.Collapsed;
                    break;
                case 3: // inform
                    UsersInput.Visibility = Visibility.Collapsed;
                    butNo.Visibility = Visibility.Collapsed;
                    break;
                default:
                    break;
            }
        }

        public void SetMsg(string _msg, string purpose, string _bot = "")
        {
            _purpose = purpose;
            MsgTextBlockMainMsg.Text = _msg;
            int msglen = _msg.Length;
            if (msglen > Width) { MaxWidth = 900; }
            MsgLabelTop.Text = purpose;
            if (_bot != "")
            {
                this.MsgLabelBot.Content = _bot;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _apath = UsersInput.Text;
            Properties.Settings.Default.FormMSG_Top = Top;
            Properties.Settings.Default.FormMSG_Left = Left;
            Properties.Settings.Default.Save();
        }
        
        public void DragWindow(object sender, MouseButtonEventArgs args)
        {
            // Watch out. Fatal error if not primary button!
            if (args.LeftButton == MouseButtonState.Pressed) { DragMove(); }
        }

        /// This button is used to be the deafult dismiss window 
        private void Quit_Click(object sender, RoutedEventArgs e)
        {
            _theResult = MessageBoxResult.Cancel;
            Close();
        }

        private void UsersInput_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string msg = "Pick the path.";
            Helpers.SetTextToASelectedFolder(sender, msg, false, _apath);
        }

        private void UsersInput_LostFocus(object sender, RoutedEventArgs e)
        {
            Helpers.EndsInBackSlash(sender);
        }

        private void ButOK_Click(object sender, RoutedEventArgs e)
        {
            _theResult = MessageBoxResult.OK;
            Close();
        }

        private void UsersInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            // adding dot mode true fakes the raw condition needed here
            Helpers.MarkTextBoxForPath(sender as TextBox, _apath, true);
        }

        private void ButNo_Click(object sender, RoutedEventArgs e)
        {
            _theResult = MessageBoxResult.No;
            Close();
        }
    }
}
