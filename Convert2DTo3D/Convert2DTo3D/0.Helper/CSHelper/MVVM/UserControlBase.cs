using System.Windows;

namespace Convert2DTo3D
{
    public class UserControlBase : UserControlValidateBase
    {
        public UserControlBase()
        {
            var _ = new Microsoft.Xaml.Behaviors.DefaultTriggerAttribute(typeof(Trigger), typeof(Microsoft.Xaml.Behaviors.TriggerBase), null);
            this.Loaded += UserControlBase_Loaded;
        }

        private void UserControlBase_Loaded(object sender, RoutedEventArgs e)
        {
            base.OnLoad(sender, e);
            ControlsUtils.Translate(this);
        }
    }
}