using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace DotNet.Business.RegReadWrite
{
    public class RegReadWrite
    {
        /// <summary>
        /// 读取路径为keypath，键名为keyname的注册表键值，缺省返回def
        /// </summary>
        /// <param name="rootkey"></param>
        /// <param name="keypath">路径</param>
        /// <param name="keyname">键名</param>
        /// <param name="rtn">默认为null</param>
        /// <returns></returns>        
        static public bool GetRegVal(RegistryKey rootkey, string keypath, string keyname, out string rtn)
        {
            rtn = "";
            try
            {
                RegistryKey key = rootkey.OpenSubKey(keypath);
                rtn = key.GetValue(keyname).ToString();
                key.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 设置路径为keypath，键名为keyname的注册表键值为keyval
        /// </summary>
        /// <param name="rootkey"></param>
        /// <param name="keypath"></param>
        /// <param name="keyname"></param>
        /// <param name="keyval"></param>
        /// <returns></returns>
        static public bool SetRegVal(RegistryKey rootkey, string keypath, string keyname, string keyval)
        {
            try
            {
                RegistryKey key = rootkey.OpenSubKey(keypath, true);
                if (key == null)
                    key = rootkey.CreateSubKey(keypath);
                key.SetValue(keyname, (object)keyval);
                key.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// 创建路径为keypath的键
        private RegistryKey CreateRegKey(RegistryKey rootkey, string keypath)
        {
            try
            {
                return rootkey.CreateSubKey(keypath);
            }
            catch
            {
                return null;
            }
        }
        /// 删除路径为keypath的子项
        private bool DelRegSubKey(RegistryKey rootkey, string keypath)
        {
            try
            {
                rootkey.DeleteSubKey(keypath);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// 删除路径为keypath的子项及其附属子项
        private bool DelRegSubKeyTree(RegistryKey rootkey, string keypath)
        {
            try
            {
                rootkey.DeleteSubKeyTree(keypath);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// 删除路径为keypath下键名为keyname的键值
        private bool DelRegKeyVal(RegistryKey rootkey, string keypath, string keyname)
        {
            try
            {
                RegistryKey key = rootkey.OpenSubKey(keypath, true);
                key.DeleteValue(keyname);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
