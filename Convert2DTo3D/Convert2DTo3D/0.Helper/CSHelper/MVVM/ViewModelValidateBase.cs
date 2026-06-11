using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;

namespace Convert2DTo3D
{
    public class ViewModelValidateBase : ValidationErrorContainer, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            try
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
            catch (InvalidOperationException ex)
            {
            }
        }

        // Called internally to notify the view that that a property has changed.
        protected void NotifyPropertyChanged(string propertyName)
        {
            if (this.PropertyChanged != null)
                this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Notify a property change that uses CallerMemberName attribute
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="backingField">Backing field of property</param>
        /// <param name="value">Value to give backing field</param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        protected virtual bool OnPropertyChanged<T>(ref T backingField, T value, [CallerMemberName] string propertyName = "")
        {
            if (EqualityComparer<T>.Default.Equals(backingField, value))
                return false;

            backingField = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        // Validate
        /// <summary>
        /// True: tức là field không bị dính lỗi validate
        /// </summary>
        /// <param name="Field"></param>
        /// <returns></returns>
        protected bool CheckValidateField(string Field)
        {
            if (lstErrors.ContainsKey(Field))
            {
                int countErr = lstErrors[Field].Count();
                if (countErr > 0) return false;
            }
            return true;
        }

        protected void RemoveValidateField(string Field)
        {
            if (lstErrors.ContainsKey(Field))
            {
                foreach (var item in lstErrors[Field])
                {
                    RemoveError(item.propertyName, item.ID);
                }
            }
        }

        /// <summary>
        /// Tính toán lại số lượng Errors
        /// </summary>
        /// <returns></returns>
        public int RefeshCountErrorFieldVisible(int totalcount, string fieldVisible, Visibility visibility)
        {
            if (visibility == Visibility.Visible)
            {
                if (CheckValidateField(fieldVisible)) RemoveValidateField(fieldVisible);
                totalcount = ErrorCount;
            }
            else
            {
                if (!CheckValidateField(fieldVisible) && lstErrors.ContainsKey(fieldVisible))
                {
                    totalcount = totalcount - 1;
                }
            }
            return totalcount;
        }

        public int RefeshEnableField(string fieldVisible, bool enableField)
        {
            if (enableField == false)
            {
                if (lstErrors.ContainsKey(fieldVisible))
                {
                    var totalcount = ErrorCount - 1;
                    return totalcount;
                }
            }
            return ErrorCount;
        }
    }
}