using System;
using System.Configuration;
using System.Data;
using System.Windows;

namespace PlayFlowMIDI
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
            {
                InputSimulator.ReleaseAllKeys();
            };

            this.DispatcherUnhandledException += (s, args) =>
            {
                InputSimulator.ReleaseAllKeys();
            };
        }
    }

}
