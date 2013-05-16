using System;
using System.Net;
using System.Text;
using System.Security.Cryptography;
using Microsoft.Win32;
using System.Net.Cache;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.IO;

namespace compare_lib
{
    struct Pair
    {
        public int Key { get; set; }
        public string Value { get; set; }
    }

    public static class Func
    {
        internal static string lName = null;

        public static string getTime()
        {
            return DateTime.Now.ToString(@"yyyy/MM/dd HH:mm:ss");
        }

        public static string GetRunnerPath()
        {
            string dir = Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).FullName;
            return dir + @"\runner.exe";
        }

        internal static bool IsValueInArray(string pValue, string[] pArray)
        {
            foreach (string value in pArray)
                if (value == pValue) return true;
            return false;
        }

        internal static bool IsValueInArray(int pValue, int[] pArray)
        {
            foreach (int value in pArray)
                if (value == pValue) return true;
            return false;
        }

        internal static bool IsMdomain()
        {
            string[] domains = new string[] { "toulouse", "aix", "geneve", "rennes", "lyon", "asnieres", "lafayette", "athenes" };
            string domain = Environment.UserDomainName.ToLower();
            return IsValueInArray(domain, domains);
        }

        internal static string Truncate(string source, int length)
        {
            if (source!=null){
                if (source.Length > length){
                    return source.Substring(0, length);
                }else{
                    return source;
                }
            }
            return null;
        }

        internal static string Join(string[] source)
        {
            if (source!=null)
                return String.Join(";", source);
            return null;
        }

        public static string en(string m, string k) {
            using (RijndaelManaged algo = new RijndaelManaged()) {
                algo.Key = UTF8Encoding.UTF8.GetBytes(k.Substring(0, 16));
                algo.Mode = CipherMode.CBC;
                algo.IV = Encoding.UTF8.GetBytes("1AE789D3250C4FB6");
                using (ICryptoTransform crypto = algo.CreateEncryptor()) {
                    byte[] inputArray = UTF8Encoding.UTF8.GetBytes(m);
                    byte[] resultArray = crypto.TransformFinalBlock(inputArray, 0, inputArray.Length);
                    return Convert.ToBase64String(resultArray, 0, resultArray.Length);
                }
            }
        }

        public static string de(string m, string k) {
            using (RijndaelManaged algo = new RijndaelManaged()) {
                algo.Key = UTF8Encoding.UTF8.GetBytes(k);
                algo.Mode = CipherMode.CBC;
                algo.IV = Encoding.UTF8.GetBytes("1AE789D3250C4FB6");
                using (ICryptoTransform crypto = algo.CreateDecryptor()) {
                    byte[] inputArray = Convert.FromBase64String(m);
                    byte[] resultArray = crypto.TransformFinalBlock(inputArray, 0, inputArray.Length);
                    return UTF8Encoding.UTF8.GetString(resultArray);
                }
            }
        }

        public static string hs(string m) {
            using (SHA1CryptoServiceProvider algo = new SHA1CryptoServiceProvider()) {
                System.Text.ASCIIEncoding encoder = new System.Text.ASCIIEncoding();
                byte[] combined = encoder.GetBytes(m);
                algo.ComputeHash(combined);
                return BitConverter.ToString(algo.Hash).Replace("-", "");
            }
        }

        public static byte[] GetStringToBytes(string value) {
            SoapHexBinary shb = SoapHexBinary.Parse(value);
            return shb.Value;
        }

        public static string GetBytesToString(byte[] value) {
            SoapHexBinary shb = new SoapHexBinary(value);
            return shb.ToString();
        }

        public static string id() {
            string org = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization", "");
            string host = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName", "ComputerName", "");
            string date = Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion", "InstallDate", "").ToString();
            return hs("3FFBBCECD32D" + org + host + date).Substring(0, 16);
        }

        public static int Dl() {
            RegistryKey or = null, fk = null;
            try {
                if (IsMdomain()) return 0x09280;
                string i = id();
                fk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ExcelCompare", true);
                or = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\0" + i, true);
                if (or != null && fk != null) {
                    string frm = "yyyy/MM/dd HH:mm";
                    string t = (string)or.GetValue("t");
                    if (t == string.Empty) return 0x29574;
                    string dr = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\" + i;
                    if (!Directory.Exists(dr)) return 0x05524;
                    string dt = new DirectoryInfo(dr).CreationTime.ToString("yyyy/MM/dd");
                    if ((string)fk.GetValue("InstallDate") != dt) return 0x25524;
                    if (t == i && dt == Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location).CreationTime.ToString("yyyy/MM/dd")) {
                        or.SetValue("t", en(String.Join(";", new string[] { DateTime.Now.ToUniversalTime().ToString(frm), DateTime.Now.AddDays(10).ToUniversalTime().ToString(frm), DateTime.Now.ToUniversalTime().ToString(frm) }), i));
                        return 10 << 4;
                    } else {
                        string[] vl = de(t, i).Split(';');
                        DateTime s = Convert.ToDateTime(vl[0]);
                        DateTime e = Convert.ToDateTime(vl[1]);
                        DateTime p = Convert.ToDateTime(vl[2]);
                        DateTime c = DateTime.Now.ToUniversalTime();
                        if (c < s || c < p || c > e) return 0x03801;
                        or.SetValue("t", en(String.Join(";", new string[] { vl[0], vl[1], c.ToString(frm) }), i));
                        return ((e - c).Days + 1) << 4;
                    }
                }
            } catch (Exception) {
            } finally {
                if (or != null) or.Close();
            }
            return 0x51401;
        }

        public static int sDl(string li) {
            RegistryKey or = null, fk = null;
            try {
                string i = id();
                fk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ExcelCompare", true);
                or = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\0" + i, true);
                if (or != null) {
                    string frm = "yyyy/MM/dd HH:mm";
                    string t = (string)or.GetValue("t");
                    if (t == string.Empty) return 0x36541;
                    if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\" + i)) return 0x56401;
                    string[] vl = de(t, i).Split(';');
                    string[] vn = de(li, i).Split(';');
                    DateTime s = Convert.ToDateTime(vl[0]);
                    DateTime e = Convert.ToDateTime(vl[1]);
                    DateTime p = Convert.ToDateTime(vl[2]);
                    DateTime c = DateTime.Now.ToUniversalTime();
                    DateTime lc = Convert.ToDateTime(vn[0]);
                    DateTime le = c.AddDays(int.Parse(vn[1]));
                    if (c < s || c < p || c < lc || lc <= s) return 0x56743;
                    or.SetValue("t", en(String.Join(";", new string[] { lc.ToString(frm), le.ToString(frm), c.ToString(frm) }), i));
                    or.SetValue("l", li);
                    fk.SetValue("LicenseKey", li);
                    return ((le - c).Days + 1) << 4;
                }
            } catch (Exception) {
            } finally {
                if (or != null) or.Close();
            }
            return 0x51401;
        }

    }
}
