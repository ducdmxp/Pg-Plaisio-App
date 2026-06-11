using System.Windows;

namespace Convert2DTo3D
{
    /// <summary>
    /// Interaction logic for StatusComponentUserControl.xaml
    /// </summary>
    public partial class StatusComponentUserControl : UserControlBase
    {
        public StatusComponentUserControl()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Window.GetWindow(this)?.Close();
        }
    }
}