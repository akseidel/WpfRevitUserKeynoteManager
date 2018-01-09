
using System;
using System.Collections.Generic;
using System.Windows;

namespace WpfRevitUserKeynoteManager
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application, ISingleInstanceApp
    {
        /// http://blogs.microsoft.co.il/arik/2010/05/28/wpf-single-instance-application/
        /// But with a variation to the methods!!
        private const string Unique = "006A3DD1-3AA6-46A2-8066-DBEA6779897C";

        void App_Startup(object sender, StartupEventArgs e)
        {
                // Regarding command line arguments. When creating the command line
                // arguments, wrap all strings that contain space with double quotes,
                // BUT, do not wrap any such string that does not have spaces. This applies
                // when there is more than one argument. Do nothing to any argument that
                // does not require double quotes to handle spaces or back slashes to 
                // handle literal quotes.

                // Application is running
                // Process command line args
                string argFolder = string.Empty;
                string argFile = string.Empty;
                //MessageBox.Show(e.Args.Length.ToString());
                if (e.Args.Length > 0)
                {
                    argFolder = e.Args[0];
                    //MessageBox.Show(argFolder);
                }
                if (e.Args.Length > 1)
                {
                    argFile = e.Args[1];
                    //MessageBox.Show(argFile);
                }

            if (SingleInstance<App>.InitializeAsFirstInstance(Unique))
            {
                // Create main application window, starting as specified
                MainWindow mainKNFWindow = new MainWindow
                {
                    Knfolder = argFolder + "/",
                    Knfile = argFile
                };
                mainKNFWindow.Show();
            } else
            {
                Current.Shutdown();
            }
        }

        #region ISingleInstanceApp Members
        public bool SignalExternalCommandLineArgs(IList<string> args)
        {
            // handle command line arguments of second instance
            // …
            return true;
        }
        #endregion
        
        private void Application_Exit(object sender, ExitEventArgs e)
        {
            // Allow single instance code to perform cleanup operations
            SingleInstance<App>.Cleanup();
        }
    }
}
