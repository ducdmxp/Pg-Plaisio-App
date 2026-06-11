using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Resources;
using System.Windows.Threading;

namespace Convert2DTo3D
{
    public class MngLanguage
    {
        private static MngLanguage _instanceB;

        private static MngLanguage Instance
        {
            get
            {
                if (_instanceB == null)
                {
                    _instanceB = new MngLanguage();
                    CurentAssembly = Assembly.GetExecutingAssembly();
                    NameSpace = CurentAssembly.FullName.Split(',')[0];
                    ResourceMng = new ResourceManager(NameSpace + ".Properties.Resources_en-US", CurentAssembly);
                }
                return _instanceB;
            }
            set => _instanceB = value;
        }

        public static ResourceManager ResourceMng { get; private set; }
        private static Assembly CurentAssembly { get; set; }
        private static string NameSpace { get; set; }

        public static void SetAssembly(Assembly ass)
        {
            CurentAssembly = ass;
            NameSpace = CurentAssembly.FullName.Split(',')[0];
        }

        public static string GetNameFromCurrentResource<T>(Expression<Func<T>> property)
        {
            var ins = Instance;
            string Name = (property.Body as MemberExpression).Member.Name;
            return ResourceMng.GetString(Name, Dispatcher.CurrentDispatcher.Thread.CurrentCulture);
        }

        public static string GetNameFromCurrentResource(string name)
        {
            return ResourceMng.GetString(name, Dispatcher.CurrentDispatcher.Thread.CurrentCulture);
        }

        public static string GetValueFromCurrentResource(string value, string groupName = "")
        {
            var entry = ResourceMng.GetResourceSet(System.Threading.Thread.CurrentThread.CurrentCulture, true, true)
                .OfType<DictionaryEntry>()
                .FirstOrDefault(e => e.Value.ToString() == value && (string.IsNullOrEmpty(groupName) ? true : e.Key.ToString().Contains(groupName)));
            return entry.Key != null ? entry.Key.ToString() : "";
        }

        public static void ChangeLanguage(string cultureName = "en-US")
        {
            var ins = Instance;
            var stringfileName = ".Properties.Resources_" + cultureName;
            ResourceMng = new ResourceManager(NameSpace + stringfileName, CurentAssembly);
            var culture = CultureInfo.GetCultureInfo(cultureName);
            Dispatcher.CurrentDispatcher.Thread.CurrentCulture = culture;
            Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = culture;
            ResourceMng = new ResourceManager(NameSpace + stringfileName, CurentAssembly);
        }
    }
}