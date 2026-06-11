using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Convert2DTo3D
{
    public enum NumberInputType
    {
        NONE = 0,
        INTEGER = 1,
        DOUBLE = 2,
    }

    public class TextBoxInteraction : DependencyObject
    {
        public static readonly DependencyProperty ValueProperty =
         DependencyProperty.RegisterAttached(
             "Value",
             typeof(double),
             typeof(TextBoxInteraction),
             new PropertyMetadata(0.0));

        public static double GetValue(TextBox textBox)
        {
            return (double)textBox.GetValue(ValueProperty);
        }

        public static void SetValue(TextBox textBox, double value)
        {
            textBox.SetValue(ValueProperty, value);
        }

        public static readonly DependencyProperty NumberInputProperty = DependencyProperty
                .RegisterAttached("NumberInput", typeof(NumberInputType), typeof(TextBoxInteraction),
                new PropertyMetadata(NumberInputType.NONE, OnNumberInputChanged));

        public static void SetNumberInput(DependencyObject d, NumberInputType use)
        {
            d.SetValue(NumberInputProperty, use);
        }

        private static void OnNumberInputChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            TextBox txt = d as TextBox;
            if (txt != null)
            {
                NumberInputType inputValue = (NumberInputType)e.NewValue;
                if (inputValue != NumberInputType.NONE)
                {
                    ContextMenu hiddenMenu = new ContextMenu
                    {
                        Visibility = Visibility.Hidden
                    };
                    txt.ContextMenu = hiddenMenu;
                    InputMethod.SetIsInputMethodEnabled(txt, false);
                    txt.GotFocus += Txt_GotFocus;
                    txt.PreviewTextInput += Txt_PreviewTextInput1;
                }
                else
                {
                    txt.PreviewTextInput -= Txt_PreviewTextInput1;
                    txt.GotFocus -= Txt_GotFocus;
                    InputMethod.SetIsInputMethodEnabled(txt, true);
                }
            }
        }

        private static void Txt_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox == null) return;
            // Đảm bảo UI đã ổn định trước khi gọi SelectAll
            textBox.Dispatcher.BeginInvoke(new Action(() =>
            {
                textBox.SelectAll();
            }), System.Windows.Threading.DispatcherPriority.Input);
        }

        private static void OnPaste(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(DataFormats.Text))
            {
                string pastedText = e.DataObject.GetData(DataFormats.Text) as string;
                if (!IsValidInput(sender, pastedText))
                {
                    e.CancelCommand();
                }
            }
            else
            {
                e.CancelCommand();
            }
        }

        private static void Txt_PreviewTextInput1(object sender, TextCompositionEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox == null) return;
            if (textBox.SelectionLength == textBox.Text.Length && textBox.Text.Length > 0) textBox.Text = "";
            var numberInput = (NumberInputType)textBox.GetValue(NumberInputProperty);
            // Check double and integer full
            if ((numberInput != NumberInputType.NONE) && e.Text == "-" && textBox.Text.Length == 0)
            {
                e.Handled = false;
                return;
            }
            // Check double
            string currentText = textBox?.Text ?? string.Empty;
            string newText = string.Empty;
            try
            {
                newText = currentText.Insert(textBox.SelectionStart, e.Text);
            }
            catch (Exception)
            {
                try
                {
                    newText = (currentText + ".").Insert(textBox.SelectionStart, e.Text);
                }
                catch (Exception)
                {
                    e.Handled = true;
                    return;
                }
            }
            // Kiểm tra tính hợp lệ
            e.Handled = !IsValidInput(sender, newText);
        }

        private static NumberInputType GetNumberInputTypeCurrent(object sender)
        {
            TextBox txt = sender as TextBox;
            if (txt == null) return NumberInputType.NONE;
            return (NumberInputType)txt.GetValue(NumberInputProperty);
        }

        private static double GetValueCurrent(object sender)
        {
            TextBox txt = sender as TextBox;
            if (txt == null) return 0.0;
            return (double)txt.GetValue(ValueProperty);
        }

        private static bool IsValidInput(object sender, string input)
        {
            var numberInputType = GetNumberInputTypeCurrent(sender);
            double valueCompare = GetValueCurrent(sender);
            if (numberInputType == NumberInputType.NONE) string.IsNullOrEmpty(input);

            if (numberInputType == NumberInputType.INTEGER)
            {
                if (int.TryParse(input, out int valueInt)) return true;
            }
            else
            {
                if (input.EndsWith("."))
                {
                    if (double.TryParse(input.TrimEnd('.'), NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                    {
                        if (numberInputType == NumberInputType.DOUBLE) return true;
                    }
                }
                if (double.TryParse(input, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedValue))
                {
                    if (numberInputType == NumberInputType.DOUBLE) return true;
                }
            }
            return string.IsNullOrEmpty(input);
        }
    }
}