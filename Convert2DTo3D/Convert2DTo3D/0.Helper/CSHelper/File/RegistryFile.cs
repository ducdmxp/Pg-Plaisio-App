using Microsoft.Win32;

namespace Convert2DTo3D
{
    public class RegistryFile
    {
        protected string m_subkey { get; set; } = @"SOFTWARE\kajima\rvt";

        public RegistryFile(string subKey)
        {
            this.m_subkey = subKey;
        }

        public string ReadFromRegistry(string keyName)
        {
            RegistryKey sk = Registry.CurrentUser.OpenSubKey(m_subkey);
            if (sk == null)
            {
                return null;
            }
            else
            {
                if (sk.GetValue(keyName.ToUpper()) != null)
                {
                    return sk.GetValue(keyName.ToUpper()).ToString();
                }
                return null;
            }
        }

        public void WriteIntoRegistry(string value, string keyname)
        {
            RegistryKey sk = Registry.CurrentUser.CreateSubKey(m_subkey);
            sk.SetValue(keyname.ToUpper(), value);
        }

        protected void DelereRegistryKey()
        {
            Registry.CurrentUser.DeleteSubKey(m_subkey);
        }

        protected string ReadFromRegistry(string keyName, string subKey)
        {
            RegistryKey sk = Registry.CurrentUser.OpenSubKey(subKey);
            if (sk == null)
            {
                return null;
            }
            else
            {
                return sk.GetValue(keyName.ToUpper()).ToString();
            }
        }

        protected void writeIntoRegistry(string keyname, string value, string subkey)
        {
            RegistryKey sk = Registry.CurrentUser.CreateSubKey(subkey);
            sk.SetValue(keyname.ToUpper(), value);
        }

        protected void DelereRegistryKey(string subkey)
        {
            Registry.CurrentUser.DeleteSubKey(subkey);
        }
    }
}