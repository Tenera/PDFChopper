using System.Windows;
using Caliburn.Micro;
using MergeSplitPdf.ViewModels;

namespace MergeSplitPdf
{
    public class AppBootstrapper : BootstrapperBase
    {
        public AppBootstrapper()
        {
            Initialize();
        }

        protected override void OnStartup(object sender, StartupEventArgs e)
        {
            DisplayRootViewFor<MainViewModel>();
        }
    }
}