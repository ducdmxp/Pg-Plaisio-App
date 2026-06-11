using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace Convert2DTo3D
{
    /*
     * Khi sử dụng TextBox để nhập số phải có 2 sự kiện:
     *    PreviewKeyDown: Ngăn không cho người dùng nhấn dấu cách
     *    TextBox_PreviewTextInput_..: Validate số nhập vào
     */

    public class UserControlValidateBase : UserControl
    {
        private List<(TextBox textbox, string propName, string errorId, string errorMessage)> _textboxes = null;

        public virtual void OnLoad(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                _textboxes = new List<(TextBox textbox, string propName, string errorId, string errorMessage)>();
                ErrorContainer = (IValidationErrorContainer)DataContext;
                AddHandler(System.Windows.Controls.Validation.ErrorEvent, new RoutedEventHandler(Handler), true);
            }
            catch (Exception)
            {
            }
        }

        public virtual void OnUnload(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                RemoveHandler(System.Windows.Controls.Validation.ErrorEvent, new RoutedEventHandler(Handler));
                foreach (var item in _textboxes)
                {
                    item.textbox.IsVisibleChanged -= Txt_IsVisibleChanged;
                    item.textbox.IsEnabledChanged -= Txt_IsEnabledChanged;
                }
            }
            catch (Exception)
            { }
        }

        internal IValidationErrorContainer ErrorContainer = null;

        // Based on exception handler from Josh Smith blog.
        public void Handler(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.ValidationErrorEventArgs args = e as System.Windows.Controls.ValidationErrorEventArgs;

            if (args.Error.RuleInError is System.Windows.Controls.ValidationRule)
            {
                if (ErrorContainer != null)
                {
                    Tracer.LogValidation("ViewBase.Handler called for ValidationRule exception.");

                    // Only want to work with validation errors that are Exceptions because the business object has already recorded the business rule violations using IDataErrorInfo.
                    BindingExpression bindingExpression = args.Error.BindingInError as System.Windows.Data.BindingExpression;
                    Debug.Assert(bindingExpression != null);

                    string propertyName = bindingExpression.ParentBinding.Path.Path;
                    DependencyObject OriginalSource = args.OriginalSource as DependencyObject;

                    // Construct the error message.
                    string errorMessage = "";
                    bool IsRemove = true;
                    ReadOnlyObservableCollection<System.Windows.Controls.ValidationError> errors = System.Windows.Controls.Validation.GetErrors(OriginalSource);
                    if (errors.Count > 0)
                    {
                        StringBuilder builder = new StringBuilder();
                        builder.Append(propertyName).Append(":");
                        System.Windows.Controls.ValidationError error = errors[errors.Count - 1];
                        {
                            if (error.Exception == null || error.Exception.InnerException == null)
                                builder.Append(error.ErrorContent.ToString());
                            else
                                builder.Append(error.Exception.InnerException.Message);
                        }
                        errorMessage = builder.ToString();
                        IsRemove = false;
                    }

                    // Add or remove the validation error to the validation error collection.
                    Debug.Assert(args.Action == ValidationErrorEventAction.Added || args.Action == ValidationErrorEventAction.Removed);
                    StringBuilder errorID = new StringBuilder();
                    errorID.Append(args.Error.RuleInError.ToString());
                    if (args.Action == ValidationErrorEventAction.Added)
                    {
                        ErrorContainer.AddError(new ValidationError(propertyName, errorID.ToString(), errorMessage));
                    }
                    else if (args.Action == ValidationErrorEventAction.Removed && IsRemove)
                    {
                        ErrorContainer.RemoveError(propertyName, errorID.ToString());
                    }
                    var txt = e.OriginalSource as TextBox;
                    if (txt != null)
                    {
                        var first = _textboxes.Where(k => k.textbox == txt).FirstOrDefault();
                        if (first.textbox != null) return;
                        txt.IsVisibleChanged += Txt_IsVisibleChanged;
                        txt.IsEnabledChanged += Txt_IsEnabledChanged;
                        _textboxes.Add((txt, propertyName, errorID.ToString(), errorMessage));
                    }
                }
            }
        }

        private void Txt_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ChangePropTextBox(sender, (bool)e.NewValue);
        }

        private void ChangePropTextBox(object sender, bool isVisible)
        {
            foreach (var item in _textboxes)
            {
                try
                {
                    TextBox txt = item.textbox as TextBox;
                    if (txt == null) continue;
                    if (txt != sender) continue;
                    if (isVisible == true && Validation.GetHasError(txt) == true)
                    {
                        ErrorContainer.AddError(new ValidationError(item.propName, item.errorId.ToString(), item.errorMessage));
                    }
                    else
                    {
                        ErrorContainer.RemoveError(item.propName, item.errorId.ToString());
                    }
                }
                catch (Exception)
                { }
            }
        }

        private void Txt_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ChangePropTextBox(sender, (bool)e.NewValue);
        }
    }
}