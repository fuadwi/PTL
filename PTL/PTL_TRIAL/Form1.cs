using Microsoft.Win32;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RYB_PTL_API;

namespace PTL_TRIAL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
        //HOW TO MAKE FEEDBACK FROM PTL ??

        private void Form1_Load(object sender, EventArgs e)
        {
            //label10.Text = "Not Connected";
            //label10.BackColor = System.Drawing.Color.Red;
            String IP = "192.168.80.253";
            RYB_PTL.RYB_PTL_Connect(IP, 6020);
            if (RYB_PTL.RYB_PTL_GetConnectState(IP) == true)
            {
                label10.Text = "Succesfully Connected to Infiniti PTL";
               label10.BackColor = System.Drawing.Color.Lime;
            }
             RYB_PTL.UserResultAvailable += RYB_PTL_UserResultAvailable;
            var appName = Process.GetCurrentProcess().ProcessName + ".exe";
            WebBrowserHelper.FixBrowserVersion();
            WebBrowserHelper.FixBrowserVersion(appName);
            webBrowser1.Url = new Uri("file:///D:/INFINITI/@OVERSEAS/EMERSON%20SINGAPORE/CamShaft%20-%20Analyze.html");


        }

        Boolean flagComplete = false;
        Boolean flagOnAction = false;
        int actionNum = 0;
        private object textBox1;
        /// <summary>
        /// 防止跨线程操作，直接赋值给TextBox文本框
        /// </summary>
        /// <param name="s"></param>
        private delegate void delegateDisplayTxtBoxScanBarcode(TextBox tx, string s);
        private void DisplayTxtBoxScanBarcode(TextBox tx, string s)
        {
            //验证是否跨线程
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new delegateDisplayTxtBoxScanBarcode(DisplayTxtBoxScanBarcode), new object[] { tx, s });
            }
            else
            {
                tb_sMsg.Text = s;
            }
        }
        private void DoPressTag(string sTagId,string sNum)
        {
            DisplayTxtBoxScanBarcode(tb_sMsg, string.Format("Press Tag:{0},num:{1}", sTagId, sNum));
            MessageBox.Show(string.Format("Press Tag:{0},num:{1}",sTagId,sNum));
        }
        private void DoScan(string barcode)
        {
            MessageBox.Show("you scan barcode is:"+barcode);
        }
        private void RYB_PTL_UserResultAvailable(RYB_PTL.RtnValueStruct rs)
        {
            switch (rs.KeyCode)
            {
                case "80":
                    //Scan barcode
                    DoScan(rs.Number);
                    break;
                case "11"://press tag model 1 color 1
                case "12":
                case "71":
                case "72":
                    //
                    DoPressTag(rs.Tagid,rs.Number);
                    break;
                case "90":// press f1 or f2

                    break;
                case "fa":
                    break;
            }
            
            String IP = "192.168.0.253";
            Console.WriteLine("Tag ID = " + rs.Tagid);
            Console.WriteLine("Number = " + rs.Number);
            Console.WriteLine("Locator = " + rs.Locator);
            Console.WriteLine("KeyCode = " + rs.KeyCode);

            if (rs.KeyCode == "80") // tag id for scanner
            {
                if (flagOnAction)
                {
                    RYB_PTL.RYB_PTL_AisleLamp(IP, "FA01", 5, 3); //FFF4 ADALAH NO ADDRESS TOWERLAMP
                    return;
                }
                if (rs.Number == "SF6090082777288")
                {
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0002", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0003", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0004", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0005", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0006", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0007", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0008", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0009", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0010", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0011", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0012", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0013", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0014", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0015", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0016", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0017", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0018", 10, 4, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0019", 10, 4, 1);
                    //label11.Text = rs.Number;
                    flagOnAction = true;

                }
                if (rs.Number == "4547648203395")
                {
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0059", 3, 3, 1);
                    RYB_PTL.RYB_PTL_DspDigit5(IP, "0060", 2, 4, 2);
                    flagOnAction = true;
                }
            }
        }



        //additional for future function

        private void SetIE8KeyforWebBrowserControl(string appName)
        {
            RegistryKey Regkey = null;
            try
            {
                // For 64 bit machine
                if (Environment.Is64BitOperatingSystem)
                    Regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Wow6432Node\\Microsoft\\Internet Explorer\\MAIN\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                else  //For 32 bit machine
                    Regkey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\\Microsoft\\Internet Explorer\\Main\\FeatureControl\\FEATURE_BROWSER_EMULATION", true);
                // If the path is not correct or
                // if the user haven't priviledges to access the registry
                if (Regkey == null)
                {
                    MessageBox.Show("Application Settings Failed - Address Not found");
                    return;
                }

                string FindAppkey = Convert.ToString(Regkey.GetValue(appName));
                  // Check if key is already present
                if (FindAppkey == "11001")
                {
                    MessageBox.Show("Required Application Settings Present");
                    Regkey.Close();
                    return;
                }
                // If a key is not present add the key, Key value 8000 (decimal)
                if (string.IsNullOrEmpty(FindAppkey))
                    Regkey.SetValue(appName, unchecked((int)0x2AF9), RegistryValueKind.DWord);
                if (Convert.ToString(Regkey.GetValue(appName)) != "11001")
                {
                    Regkey.SetValue(appName, unchecked((int)0x2AF9), RegistryValueKind.DWord);
                                    }
                // Check for the key after adding
                FindAppkey = Convert.ToString(Regkey.GetValue(appName));
                                if (FindAppkey == "11001")
                    MessageBox.Show("Application Settings Applied Successfully");
                else
                    MessageBox.Show("Application Settings Failed, Ref: " + FindAppkey);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Application Settings Failed");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // Close the Registry
                if (Regkey != null)
                    Regkey.Close();
            }
        }
        public class WebBrowserHelper
        {


            public static int GetEmbVersion()
            {
                int ieVer = GetBrowserVersion();

                if (ieVer > 9)
                    return ieVer * 1000 + 1;

                if (ieVer > 7)
                    return ieVer * 1111;

                return 7000;
            } // End Function GetEmbVersion

            public static void FixBrowserVersion()
            {
                string appName = System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location);
                FixBrowserVersion(appName);
            }

            public static void FixBrowserVersion(string appName)
            {
                FixBrowserVersion(appName, GetEmbVersion());
            } // End Sub FixBrowserVersion

            // FixBrowserVersion("<YourAppName>", 9000);
            public static void FixBrowserVersion(string appName, int ieVer)
            {
                FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName + ".exe", ieVer);
                FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName + ".exe", ieVer);
                FixBrowserVersion_Internal("HKEY_LOCAL_MACHINE", appName + ".vshost.exe", ieVer);
                FixBrowserVersion_Internal("HKEY_CURRENT_USER", appName + ".vshost.exe", ieVer);
            } // End Sub FixBrowserVersion 

            private static void FixBrowserVersion_Internal(string root, string appName, int ieVer)
            {
                try
                {
                    //For 64 bit Machine 
                    if (Environment.Is64BitOperatingSystem)
                        Microsoft.Win32.Registry.SetValue(root + @"\Software\Wow6432Node\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", appName, ieVer);
                    else  //For 32 bit Machine 
                        Microsoft.Win32.Registry.SetValue(root + @"\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", appName, ieVer);


                }
                catch (Exception)
                {
                    // some config will hit access rights exceptions
                    // this is why we try with both LOCAL_MACHINE and CURRENT_USER
                }
            } // End Sub FixBrowserVersion_Internal 

            public static int GetBrowserVersion()
            {
                // string strKeyPath = @"HKLM\SOFTWARE\Microsoft\Internet Explorer";
                string strKeyPath = @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer";
                string[] ls = new string[] { "svcVersion", "svcUpdateVersion", "Version", "W2kVersion" };

                int maxVer = 0;
                for (int i = 0; i < ls.Length; ++i)
                {
                    object objVal = Microsoft.Win32.Registry.GetValue(strKeyPath, ls[i], "0");
                    string strVal = System.Convert.ToString(objVal);
                    if (strVal != null)
                    {
                        int iPos = strVal.IndexOf('.');
                        if (iPos > 0)
                            strVal = strVal.Substring(0, iPos);

                        int res = 0;
                        if (int.TryParse(strVal, out res))
                            maxVer = Math.Max(maxVer, res);
                    } // End if (strVal != null)

                } // Next i

                return maxVer;
            } // End Function GetBrowserVersion 


        }

        private void button1_Click(object sender, EventArgs e)
        {
            string IP = "192.168.0.253";
            RYB_PTL.RYB_PTL_AisleLamp(IP, "FA01", 7, 1);
            RYB_PTL.RYB_PTL_DspDigit5(IP, "0004", Convert.ToInt32(textBo1.Text), 2, 3);
            RYB_PTL.RYB_PTL_DspDigit5(IP, "0005", Convert.ToInt32(textBo2.Text), 1, 1);
            RYB_PTL.RYB_PTL_DspDigit5(IP, "0006", Convert.ToInt32(textBo3.Text), 3, 2);
            RYB_PTL.RYB_PTL_DspDigit5(IP, "0007", Convert.ToInt32(textBo4.Text), 4, 3);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            String IP = "192.168.0.253";
            RYB_PTL.RYB_PTL_CloseDigit5(IP, "AAAA");
            
         //   RYB_PTL.RYB_PTL_AisleLamp(IP, "FFF4", 7, 0);
           // RYB_PTL.RYB_PTL_AisleLamp(IP, "FFF4", 6, 0);
            RYB_PTL.RYB_PTL_Disconnect(IP);
            RYB_PTL.UserResultAvailable -= RYB_PTL_UserResultAvailable;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String IP = "192.168.0.253";
            RYB_PTL.RYB_PTL_Connect(IP, 6020);
            if (RYB_PTL.RYB_PTL_GetConnectState(IP) == true)
            {
                label10.Text = "Succesfully Connected to Infiniti PTL";
                label10.BackColor = System.Drawing.Color.Lime;
                RYB_PTL.RYB_PTL_CloseDigit5(IP,"AAAA");
            }
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBo1_TextChanged(object sender, EventArgs e)
        {

        }
    }

}
