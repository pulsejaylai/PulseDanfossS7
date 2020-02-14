using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using System.IO.Ports;
using System.IO;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for Equipment.xaml
    /// </summary>
    public partial class Equipment : Window
    {
        public Equipment()
        {
            InitializeComponent();
        }
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);
        [DllImport("kernel32")]
        public static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport(@"D:\GPIBDLL.dll", EntryPoint = "Gpiblist", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern IntPtr equlist();
        public class MeaType
        {
            public int ID { get; set; }
            public string add { get; set; }
        }
        public class MeaType2
        {
            public int IDW { get; set; }
            public string addw { get; set; }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            FileStream fs;
            StringBuilder temp = new StringBuilder(500);
            int portx = 1,index=0;
            string[] ports = SerialPort.GetPortNames();
            string[] equ;
               IntPtr intPtr = equlist();
               string str = Marshal.PtrToStringAnsi(intPtr);
               // Delay(100);
               IntPtr intPtr2 = equlist();

               string str2 = Marshal.PtrToStringAnsi(intPtr2);
               equ = str2.Split(',');
               List<MeaType> list = new List<MeaType>();
               foreach (string port in equ)
               { list.Add(new MeaType { ID = portx, add = port }); portx++; }   
          
            motoradd.ItemsSource = list;
            keiadd.ItemsSource = list;
            pow1add.ItemsSource = list;
            pow2add.ItemsSource = list;
            List<MeaType2> list2 = new List<MeaType2>();
            list2.Add(new MeaType2 { IDW = 0, addw = "Dev2/ai0" });
            list2.Add(new MeaType2 { IDW = 1, addw = "Dev2/ai1" });
            list2.Add(new MeaType2 { IDW = 2, addw = "Dev2/ai2" });
            list2.Add(new MeaType2 { IDW = 3, addw = "Dev2/ai3" });
            list2.Add(new MeaType2 { IDW = 4, addw = "Dev2/ai4" });
            nc1add.ItemsSource = list2;
            no2add.ItemsSource = list2;
            nc3add.ItemsSource = list2;
            no4add.ItemsSource = list2;
            ledadd.ItemsSource = list2;
            if (!File.Exists("D:\\HardInfo.ini"))
            {

                fs = new FileStream("D:\\HardInfo.ini", FileMode.Create);
                StreamWriter checksn = new StreamWriter(fs);
                checksn.Flush();
                checksn.Close();
                fs.Close();
            }
else
            {
                  GetPrivateProfileString("HardInfo", "StepMotor", "0", temp, 500, "D:\\HardInfo.ini");
                int xx = 0;
                xx = list.FindIndex(delegate (MeaType user) { return user.add == temp.ToString(); });               
                index = motoradd.Items.IndexOf(list[xx]);
               // MessageBox.Show(list[1].add.ToString());
                motoradd.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "Keithley", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list.FindIndex(delegate (MeaType user) { return user.add == temp.ToString(); });
                index = keiadd.Items.IndexOf(list[xx]);
                // MessageBox.Show(list[1].add.ToString());
                keiadd.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "Pow1", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list.FindIndex(delegate (MeaType user) { return user.add == temp.ToString(); });
                index = pow1add.Items.IndexOf(list[xx]);
                // MessageBox.Show(list[1].add.ToString());
                pow1add.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "Pow2", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list.FindIndex(delegate (MeaType user) { return user.add == temp.ToString(); });
                index = pow2add.Items.IndexOf(list[xx]);
                // MessageBox.Show(list[1].add.ToString());
                pow2add.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "NC1", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list2.FindIndex(delegate (MeaType2 user2) { return user2.addw == temp.ToString(); });
                index = nc1add.Items.IndexOf(list2[xx]);
                // MessageBox.Show(list[1].add.ToString());
                nc1add.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "NO2", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list2.FindIndex(delegate (MeaType2 user2) { return user2.addw == temp.ToString(); });
                index = no2add.Items.IndexOf(list2[xx]);
                // MessageBox.Show(list[1].add.ToString());
                no2add.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "NC3", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list2.FindIndex(delegate (MeaType2 user2) { return user2.addw == temp.ToString(); });
                index = nc3add.Items.IndexOf(list2[xx]);
                // MessageBox.Show(list[1].add.ToString());
                nc3add.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "NO4", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list2.FindIndex(delegate (MeaType2 user2) { return user2.addw == temp.ToString(); });
                index = no4add.Items.IndexOf(list2[xx]);
                // MessageBox.Show(list[1].add.ToString());
                no4add.SelectedIndex = index;
                xx = 0;
                GetPrivateProfileString("HardInfo", "LED", "0", temp, 500, "D:\\HardInfo.ini");
                xx = list2.FindIndex(delegate (MeaType2 user2) { return user2.addw == temp.ToString(); });
                index = ledadd.Items.IndexOf(list2[xx]);
                // MessageBox.Show(list[1].add.ToString());
                ledadd.SelectedIndex = index;

            }


        }

        private void Writeini_Click(object sender, RoutedEventArgs e)
        {
            WritePrivateProfileString("Hardinfo", "StepMotor", motoradd.Text.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "Keithley", keiadd.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "Pow1", pow1add.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "Pow2", pow2add.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "NC1", nc1add.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "NO2", no2add.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "NC3", nc3add.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "NO4", no4add.ToString(), "D:\\HardInfo.ini");
            WritePrivateProfileString("Hardinfo", "LED", ledadd.ToString(), "D:\\HardInfo.ini");

        }




    }
}
