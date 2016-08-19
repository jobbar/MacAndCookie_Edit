using Microsoft.Win32;
using NETCONLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MacAndCookie_Edit
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnClearCookie_Click(object sender, RoutedEventArgs e)
        {
            //Temporary Internet Files  （Internet临时文件）
            RunCmd("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8");

            //Cookies
            RunCmd("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2");

            //Passwords (密码）
            RunCmd("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32");

            //全部删除
            RunCmd("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351");

        }
        private void RunCmd(string cmd)
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.FileName = "cmd.exe";
            // 关闭Shell的使用
            p.StartInfo.UseShellExecute = false;
            // 重定向标准输入
            p.StartInfo.RedirectStandardInput = true;
            // 重定向标准输出
            p.StartInfo.RedirectStandardOutput = true;
            //重定向错误输出
            p.StartInfo.RedirectStandardError = true;
            p.StartInfo.CreateNoWindow = true;
            p.Start();
            p.StandardInput.WriteLine(cmd);
            p.StandardInput.WriteLine("exit");
        }


        private void btnEditMac_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }


    public class MACHelper
    {
        [DllImport("wininet.dll")]
        private extern static bool InternetGetConnectedState(int Description, int ReservedValue);
        /// <summary>
        /// 是否能连接上Internet
        /// </summary>
        /// <returns></returns>
        public bool IsConnectedToInternet()
        {
            int Desc = 0;
            return InternetGetConnectedState(Desc, 0);
        }
        /// <summary>
        /// 获取MAC地址
        /// </summary>
        public string GetMACAddress()
        {
            //得到 MAC的注册表键
            RegistryKey macRegistry = Registry.LocalMachine.OpenSubKey("SYSTEM").OpenSubKey("CurrentControlSet").OpenSubKey("Control")
                .OpenSubKey("Class").OpenSubKey("{4D36E972-E325-11CE-BFC1-08002bE10318}");
            IList<string> list = macRegistry.GetSubKeyNames().ToList();
            IPGlobalProperties computerProperties = IPGlobalProperties.GetIPGlobalProperties();
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            var adapter = nics.First(o => o.Name == "本地连接");
            if (adapter == null)
                return null;
            return string.Empty;
        }

        /// <summary>
        /// 设置MAC地址
        /// </summary>
        /// <param name="newMac"></param>
        public void SetMACAddress(string newMac)
        {
            string macAddress;
            string index = GetAdapterIndex(out macAddress);
            if (index == null)
                return;
            //得到 MAC的注册表键
            RegistryKey macRegistry = Registry.LocalMachine.OpenSubKey("SYSTEM").OpenSubKey("CurrentControlSet").OpenSubKey("Control")
                .OpenSubKey("Class").OpenSubKey("{4D36E972-E325-11CE-BFC1-08002bE10318}").OpenSubKey(index, true);
            if (string.IsNullOrEmpty(newMac))
            {
                macRegistry.DeleteValue("NetworkAddress");
            }
            else
            {
                macRegistry.SetValue("NetworkAddress", newMac);
                macRegistry.OpenSubKey("Ndi", true).OpenSubKey("params", true).OpenSubKey("NetworkAddress", true).SetValue("Default", newMac);
                macRegistry.OpenSubKey("Ndi", true).OpenSubKey("params", true).OpenSubKey("NetworkAddress", true).SetValue("ParamDesc", "Network Address");
            }
            Thread oThread = new Thread(new ThreadStart(ReConnect));//new Thread to ReConnect
            oThread.Start();
        }
        /// <summary>
        /// 重设MAC地址
        /// </summary>
        public void ResetMACAddress()
        {
            SetMACAddress(string.Empty);
        }
        /// <summary>
        /// 重新连接
        /// </summary>
        private void ReConnect()
        {
            NetSharingManagerClass netSharingMgr = new NetSharingManagerClass();
            INetSharingEveryConnectionCollection connections = netSharingMgr.EnumEveryConnection;
            foreach (INetConnection connection in connections)
            {
                INetConnectionProps connProps = netSharingMgr.get_NetConnectionProps(connection);
                if (connProps.MediaType == tagNETCON_MEDIATYPE.NCM_LAN)
                {
                    connection.Disconnect(); //禁用网络
                    connection.Connect();    //启用网络
                }
            }
        }
        /// <summary>
        /// 生成随机MAC地址
        /// </summary>
        /// <returns></returns>
        public string CreateNewMacAddress()
        {
            //return "0016D3B5C493";
            int min = 0;
            int max = 16;
            Random ro = new Random();
            var sn = string.Format("{0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}",
               ro.Next(min, max).ToString("x"),//0
               ro.Next(min, max).ToString("x"),//
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),//5
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),
               ro.Next(min, max).ToString("x"),//10
               ro.Next(min, max).ToString("x")
                ).ToUpper();
            return sn;
        }
        /// <summary>
        /// 得到Mac地址及注册表对应Index
        /// </summary>
        /// <param name="macAddress"></param>
        /// <returns></returns>
        public string GetAdapterIndex(out string macAddress)
        {
            ManagementClass oMClass = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection colMObj = oMClass.GetInstances();
            macAddress = string.Empty;
            int indexString = 1;
            foreach (ManagementObject objMO in colMObj)
            {
                indexString++;
                if (objMO["MacAddress"] != null && (bool)objMO["IPEnabled"] == true)
                {
                    macAddress = objMO["MacAddress"].ToString().Replace(":", "");
                    break;
                }
            }
            if (macAddress == string.Empty)
                return null;
            else
                return indexString.ToString().PadLeft(4, '0');
        }
        #region Temp


        public void noting()
        {
            //ManagementClass oMClass = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementClass oMClass = new ManagementClass("Win32_NetworkAdapter");
            ManagementObjectCollection colMObj = oMClass.GetInstances();
            foreach (ManagementObject objMO in colMObj)
            {
                if (objMO["MacAddress"] != null)
                {
                    if (objMO["Name"] != null)
                    {
                        //objMO.InvokeMethod("Reset", null);
                        objMO.InvokeMethod("Disable", null);//Vista only
                        objMO.InvokeMethod("Enable", null);//Vista only
                    }
                    //if ((bool)objMO["IPEnabled"] == true)
                    //{
                    //    //Console.WriteLine(objMO["MacAddress"].ToString());
                    //    //objMO.SetPropertyValue("MacAddress", CreateNewMacAddress());
                    //    //objMO["MacAddress"] = CreateNewMacAddress();
                    //    //objMO.InvokeMethod("Disable", null);
                    //    //objMO.InvokeMethod("Enable", null);
                    //    //objMO.Path.ReleaseDHCPLease();
                    //    var iObj = objMO.GetMethodParameters("EnableDHCP");
                    //    var oObj = objMO.InvokeMethod("ReleaseDHCPLease", null, null);
                    //    Thread.Sleep(100);
                    //    objMO.InvokeMethod("RenewDHCPLease", null, null);
                    //}
                }
            }
        }
        public void no()
        {
            Shell32.Folder networkConnectionsFolder = GetNetworkConnectionsFolder();
            if (networkConnectionsFolder == null)
            {
                Console.WriteLine("Network connections folder not found.");
                return;
            }
            Shell32.FolderItem2 networkConnection = GetNetworkConnection(networkConnectionsFolder, string.Empty);
            if (networkConnection == null)
            {
                Console.WriteLine("Network connection not found.");
                return;
            }
            Shell32.FolderItemVerb verb;
            try
            {
                IsNetworkConnectionEnabled(networkConnection, out verb);
                verb.DoIt();
                Thread.Sleep(1000);
                IsNetworkConnectionEnabled(networkConnection, out verb);
                verb.DoIt();
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        /// <summary>
        /// Gets the Network Connections folder in the control panel.
        /// </summary>
        /// <returns>The Folder for the Network Connections folder, or null if it was not found.</returns>
        static Shell32.Folder GetNetworkConnectionsFolder()
        {
            Shell32.Shell sh = new Shell32.Shell();
            Shell32.Folder controlPanel = sh.NameSpace(3); // Control panel
            Shell32.FolderItems items = controlPanel.Items();
            foreach (Shell32.FolderItem item in items)
            {
                if (item.Name == "网络连接")
                    return (Shell32.Folder)item.GetFolder;
            }
            return null;
        }
        /// <summary>
        /// Gets the network connection with the specified name from the specified shell folder.
        /// </summary>
        /// <param name="networkConnectionsFolder">The Network Connections folder.</param>
        /// <param name="connectionName">The name of the network connection.</param>
        /// <returns>The FolderItem for the network connection, or null if it was not found.</returns>
        static Shell32.FolderItem2 GetNetworkConnection(Shell32.Folder networkConnectionsFolder, string connectionName)
        {
            Shell32.FolderItems items = networkConnectionsFolder.Items();
            foreach (Shell32.FolderItem2 item in items)
            {
                if (item.Name == "本地连接")
                {
                    return item;
                }
            }
            return null;
        }
        /// <summary>
        /// Gets whether or not the network connection is enabled and the command to enable/disable it.
        /// </summary>
        /// <param name="networkConnection">The network connection to check.</param>
        /// <param name="enableDisableVerb">On return, receives the verb used to enable or disable the connection.</param>
        /// <returns>True if the connection is enabled, false if it is disabled.</returns>
        static bool IsNetworkConnectionEnabled(Shell32.FolderItem2 networkConnection, out Shell32.FolderItemVerb enableDisableVerb)
        {
            Shell32.FolderItemVerbs verbs = networkConnection.Verbs();
            foreach (Shell32.FolderItemVerb verb in verbs)
            {
                if (verb.Name == "启用(&A)")
                {
                    enableDisableVerb = verb;
                    return false;
                }
                else if (verb.Name == "停用(&B)")
                {
                    enableDisableVerb = verb;
                    return true;
                }
            }
            throw new ArgumentException("No enable or disable verb found.");
        }
        #endregion
    }
}
