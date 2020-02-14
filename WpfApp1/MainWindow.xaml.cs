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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using System.Windows.Threading;
using System.Security.Permissions;
using System.ComponentModel;
using System.Text.RegularExpressions;
using Ivi.Visa.Interop;
using System.Threading;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Controls.Primitives;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // ErrBox.ItemsSource = Setcontent(string c);
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.DoWork += DoWork_Handler;
            bgWorker.ProgressChanged += ProgressChanged_Handler;
            bgWorker.RunWorkerCompleted += RunWorkerCompleted_Handler;

        }
        BackgroundWorker bgWorker = new BackgroundWorker();
        BackgroundWorker bgWorker2 = new BackgroundWorker();
        string ff, ff1, Save_Data, keithleyadd, snno = "0000", pow1add = "0", pow2add = "0", NC1add, NO2add, NC3add, NO4add, LEDadd, DCRADD, CanEnable = "0", CoilNumber, snsample = "", rev, pcbrev, pdays, smodel, savemodel, path2, savepath,errcode="",dishex;
        SqlConnection con = new SqlConnection();
        SqlCommand cmd = new SqlCommand();
        int single = 0;
        int x_p = 0, x_s = 0;
        ObservableCollection<Testdata> testdata = new ObservableCollection<Testdata>();
        // ErrBox.ItemsSource=Setcontent();
        void frm_TransfEvent(string value)
        {
            snsample = value;
        }
        void frm_TransintfEvent(int valuet)
        {
            samplelength = valuet;
        }
        static string getrev(String num, String result)
        {
            char[] numch = num.ToCharArray();
            int x, revposition, revposition2;
            x = 0;
            revposition = 0;
            revposition2 = 0;
            for (int i = 0; i < num.Length; i++)
            {
                if (numch[i].ToString() == ":")
                {
                    x++;
                    if (x == 1)
                    {
                        revposition = i;
                        // break;
                    }
                    if (x == 2)
                    {
                        revposition2 = i;
                        // break;
                    }

                }

            }
            result = num.Substring(revposition + 1, revposition2 - revposition - 1);
            // MessageBox.Show(result);
            return result;
        }
        static string getpcbrev(String num, String result)
        {
            char[] numch = num.ToCharArray();
            int x, revposition, revposition2, count;
            x = 0;
            count = 0;
            revposition = 0;
            revposition2 = 0;
            for (int i2 = 0; i2 < num.Length; i2++)
            {
                if (numch[i2].ToString() == ":")
                {
                    count++;
                }
            }
            if (count > 2)
            {
                for (int i = 0; i < num.Length; i++)
                {
                    if (numch[i].ToString() == ":")
                    {
                        x++;

                        if (x == 2)
                        {
                            revposition = i;
                            // break;
                        }
                        if (x == 3)
                        {
                            revposition2 = i;
                            // break;
                        }


                    }
                }
            }

            if (count > 2)
            { result = num.Substring(revposition + 1, revposition2 - revposition - 1); }
            else
            { result = ""; }
            return result;
        }
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (CheckBox1.IsChecked == true)
            {
                SampleSN smsninput=new SampleSN();
                Regex regnum = new Regex("^[0~9]");
                smsninput.TransfEvent += frm_TransfEvent;
                smsninput.TransintfEvent += frm_TransintfEvent;
                smsninput.ShowDialog();

                int sindex;
                string revx, revr;
                if (samplelength != 0)
                { test.IsEnabled = true; }
                revr = "";
                sindex = snsample.IndexOf(":");
                smodel = snsample.Substring(0, sindex);
                //  MessageBox.Show(smodel);
                revx = getrev(snsample, revr);
                //  MessageBox.Show(rev);  

                if (regnum.IsMatch(revx.Substring(0, 1)))
                {
                    // MessageBox.Show("Num");
                    rev = revx.Substring(0, 2);
                    pdays = revx.Substring(2, 4);
                    // MessageBox.Show(rev);
                }
                else
                {
                    rev = revx.Substring(0, 1);
                    pdays = revx.Substring(1, 4);
                    // MessageBox.Show(rev);
                }
                revr = "";
                pcbrev = getpcbrev(snsample, revr);





            }
            test.Focus();
        }
        [DllImport("kernel32.dll")]
        public static extern uint GetTickCount();
        public static class DispatcherHelper
        {
            [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
            public static void DoEvents()
            {
                DispatcherFrame frame = new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrames), frame);
                try { Dispatcher.PushFrame(frame); }
                catch (InvalidOperationException) { }
            }
            private static object ExitFrames(object frame)
            {
                ((DispatcherFrame)frame).Continue = false;
                return null;
            }
        }
        public static void Delay(uint ms)
        {
            uint start = GetTickCount();
            while (GetTickCount() - start < ms)
            {
                DispatcherHelper.DoEvents();
            }
        }
        private string canreceive2(out Ecan.CAN_OBJ[] resultobj, int count,UInt32 canchannl)
        {
            string canresult = "";
            Ecan.CAN_ERR_INFO errinfo;
            int workStationCount = count;
            int size = Marshal.SizeOf(typeof(Ecan.CAN_OBJ));
            IntPtr infosIntptr = Marshal.AllocHGlobal(size * workStationCount);
            resultobj = new Ecan.CAN_OBJ[workStationCount];

            //    MessageBox.Show(count.ToString());

            Delay(40);
            if (Ecan.Receive2(4, 0, canchannl, infosIntptr, (ushort)count, 10) == Ecan.ECANStatus.STATUS_OK) {/* MessageBox.Show(resultobj.ID.ToString("X"));*/ Delay(30); }
            else
            {
                Ecan.ReadErrInfo(4, 0, canchannl, out errinfo);
                canresult = "ReciveCanbus ErrCode:" + errinfo.ErrCode.ToString("X");
                //   for (int i = 0; i < 8; i++) { resultobj.data[i] = 0xFF; }
                // resultobj.ID = 268435455;
            }
            for (int inkIndex = 0; inkIndex < workStationCount; inkIndex++)
            {
                IntPtr ptr = (IntPtr)((UInt32)infosIntptr + inkIndex * size);
                resultobj[inkIndex] = (Ecan.CAN_OBJ)Marshal.PtrToStructure(ptr, typeof(Ecan.CAN_OBJ));
            }

            //   MessageBox.Show(resultobj[0].ID.ToString());
            // MessageBox.Show(resultobj[1].ID.ToString());
            Ecan.ClearCanbuf(4, 0, canchannl);
            Delay(30);
            return canresult;
        }
        private string canreceive(out Ecan.CAN_OBJ resultobj)
        {
            string canresult = "";
            Ecan.CAN_ERR_INFO errinfo;
            resultobj = new Ecan.CAN_OBJ();



            Delay(40);
            if (Ecan.Receive(4, 0, 0, out resultobj, (ushort)1, 10) == Ecan.ECANStatus.STATUS_OK) {/* MessageBox.Show(resultobj.ID.ToString("X"));*/ Delay(30); }
            else
            {
                Ecan.ReadErrInfo(4, 0, 0, out errinfo);
                canresult = "ReciveCanbus ErrCode:" + errinfo.ErrCode.ToString("X");
                for (int i = 0; i < 8; i++) { resultobj.data[i] = 0xFF; }
                resultobj.ID = 268435455;
            }


            Ecan.ClearCanbuf(4, 0, 0);
            return canresult;
        }

        private string cansend2(Ecan.CAN_OBJ caninfo, out Ecan.CAN_OBJ[] resultobj, int count,UInt32 canchannl)
        {
            string canresult = "";
            Ecan.CAN_ERR_INFO errinfo;
            int workStationCount = count;
            int size = Marshal.SizeOf(typeof(Ecan.CAN_OBJ));
            IntPtr infosIntptr = Marshal.AllocHGlobal(size * workStationCount);
            resultobj = new Ecan.CAN_OBJ[workStationCount];
            //resultobj = new Ecan.CAN_OBJ[count];
            try
            {
                if (Ecan.Transmit(4, 0, canchannl, ref caninfo, (ushort)1) != Ecan.ECANStatus.STATUS_OK)
                {
                    Ecan.ReadErrInfo(4, 0, canchannl, out errinfo);
                    canresult = "SendCanBus ErrCode:" + errinfo.ErrCode.ToString("X");
                    // for (int i = 0; i < 8; i++) { resultobj.data[i] = 0xFF; }
                    // resultobj.ID = 268435455;

                }
                else
                {
                    // if (canmodel == "HW Status") { Delay(13000); }
                    Delay(40);
                    if (Ecan.Receive2(4, 0, canchannl, infosIntptr, (ushort)count, 10) == Ecan.ECANStatus.STATUS_OK) {/* MessageBox.Show(resultobj.ID.ToString("X"));*/ Delay(30); }
                    else
                    {
                        Ecan.ReadErrInfo(4, 0, canchannl, out errinfo);
                        canresult = "ReciveCanbus ErrCode:" + errinfo.ErrCode.ToString("X");
                        //  for (int i = 0; i < 8; i++) { resultobj.data[i] = 0xFF; }
                        // resultobj.ID = 268435455;
                    }
                    for (int inkIndex = 0; inkIndex < workStationCount; inkIndex++)
                    {
                        IntPtr ptr = (IntPtr)((UInt32)infosIntptr + inkIndex * size);
                        resultobj[inkIndex] = (Ecan.CAN_OBJ)Marshal.PtrToStructure(ptr, typeof(Ecan.CAN_OBJ));
                    }

                }

            }
            catch (Exception ee)
            {
                errcontent2 = ee.Message;
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
            }
            Ecan.ClearCanbuf(4, 0, 0);







            return canresult;
        }




        private string cansend(Ecan.CAN_OBJ caninfo, out Ecan.CAN_OBJ resultobj)
        {
            string canresult = "";
            Ecan.CAN_ERR_INFO errinfo;
            resultobj = new Ecan.CAN_OBJ();
            try
            {
                if (Ecan.Transmit(4, 0, 0, ref caninfo, (ushort)1) != Ecan.ECANStatus.STATUS_OK)
                {
                    Ecan.ReadErrInfo(4, 0, 0, out errinfo);
                    canresult = "SendCanBus ErrCode:" + errinfo.ErrCode.ToString("X");
                    for (int i = 0; i < 8; i++) { resultobj.data[i] = 0xFF; }
                    resultobj.ID = 268435455;

                }
                else
                {
                    if (canmodel == "HW Status") { Delay(13000); }
                    Delay(40);
                    if (Ecan.Receive(4, 0, 0, out resultobj, (ushort)1, 10) == Ecan.ECANStatus.STATUS_OK) {/* MessageBox.Show(resultobj.ID.ToString("X"));*/ Delay(30); }
                    else
                    {
                        Ecan.ReadErrInfo(4, 0, 0, out errinfo);
                        canresult = "ReciveCanbus ErrCode:" + errinfo.ErrCode.ToString("X");
                        for (int i = 0; i < 8; i++) { resultobj.data[i] = 0xFF; }
                        resultobj.ID = 268435455;
                    }

                }
            }
            catch (Exception ee)
            {
                errcontent2 = ee.Message;
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
               
            }
            Ecan.ClearCanbuf(4, 0, 0);
            return canresult;
        }
        [DllImport(@"D:\GPIBDLL.dll", EntryPoint = "Gpread", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern IntPtr gpread([MarshalAs(UnmanagedType.LPStr)]string add, [MarshalAs(UnmanagedType.LPStr)]string cmd);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "StopTask", CallingConvention = CallingConvention.Cdecl)]
        public static extern Int32 StopTask(IntPtr taskhandel);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "ClearTask", CallingConvention = CallingConvention.Cdecl)]
        public static extern Int32 ClearTask(IntPtr taskhandel);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "StartTask", CallingConvention = CallingConvention.Cdecl)]
        public static extern Int32 StartTask(IntPtr taskhandel);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "ConfigSampleClk", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern Int32 ConfigSampleClk(IntPtr taskhandel, double rate, Int32 activeEdge, Int32 sampleMode, UInt64 sampsPerChanToAcquire);
        public delegate string MEAS_FRE(string add2100);
        MEAS_FRE measres = (add2100) =>
        {
            string result = "";
            IntPtr intPtr = gpread(add2100, "MEAS:FRES? 1,100");
            result = Marshal.PtrToStringAnsi(intPtr);
            if (result == "") { result = "-1E+2" + "\n"; }
            return result;
        };

        public delegate string Outport(string Chan, UInt32 Data);
        Outport outport = (Chan, Data) =>
        {
            string message = "";
            UInt32 handel;
            IntPtr pA, intPtr;
            Int32 errcode;
            Int32 writenum;
            handel = CreatTask();
            pA = new IntPtr(handel);
            errcode = ConfigDOChann(pA, Chan, "");
            errcode = StartTask(pA);
            errcode = Writeport(pA, 1, 1, 5.0, Data, out writenum);
            if (errcode != 0)
            {
                intPtr = GetErr(errcode);
                message = Marshal.PtrToStringAnsi(intPtr);

            }
            errcode = StopTask(pA);
            errcode = ClearTask(pA);
            return message;
        };
        public delegate double[] MeasVol2(string Chan, int count);
        MeasVol2 measvol2 = (Chan, count) =>
        {
            int i;
            Int32 errcode, readnum;
            IntPtr intPtr, pA;
            UInt32 handel;
            double[] data;
            data = new double[count];
            handel = CreatTask();
            pA = new IntPtr(handel);
            ConfigChann(pA, Chan, "", 10083, -10.0, 10.0, 10348); ConfigSampleClk(pA, 10000.0, 10280, 10178, 1000); StartTask(pA);
            errcode = GetDCVol(pA, 1000, 10, data, out readnum);
            StopTask(pA);
            ClearTask(pA);


            return data;
        };

        public delegate double MeasVol(string Chan);
        MeasVol measvol = (Chan) =>
        {
            double result = -1000.0, zong = 0.0;
            int i;
            Int32 errcode, readnum;
            IntPtr intPtr, pA;
            UInt32 handel;
            double[] data;
            data = new double[1000];
            handel = CreatTask();
            pA = new IntPtr(handel);
            ConfigChann(pA, Chan, "", 10083, -10.0, 10.0, 10348); ConfigSampleClk(pA, 10000.0, 10280, 10178, 1000); StartTask(pA);
            errcode = GetDCVol(pA, 1000, 10, data, out readnum);
            if (errcode == 0)
            {
                for (i = 0; i < 1000; i++) { zong = zong + data[i]; }
                result = zong / 1000;
            }
            StopTask(pA);
            ClearTask(pA);


            return result;
        };





        public delegate string Send(SerialPort sp1, string common);
        Send sendcomm = (sp1, common) =>
        {
            // SerialPort sp1;
            //sp1 = new SerialPort("com1");
            string errmessage = "";
            byte[] ee2 = new byte[common.Length / 2];
            for (int i = 0; i < ee2.Length; i++)
            {
                ee2[i] = Convert.ToByte(common.Substring(i * 2, 2), 16);
            }
            Delay(50);
            //    sp.Open();
            try
            {
                sp1.Write(ee2, 0, 9);
            }
            catch (Exception ee)
            {
                //MessageBox.Show(ee.Message);
                errmessage = ee.ToString();
            }
            return errmessage;
        };

        FileStream passcheck;


        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            #region //判断系统是否已启动

            System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName("WpfApp1");//获取指定的进程名   
            if (myProcesses.Length > 1) //如果可以获取到知道的进程名则说明已经启动
            {
                MessageBox.Show("程序已启动！");
                System.Environment.Exit(0);              //关闭系统
            }
            //

            #endregion
            FileStream passtext;
            string checkp;
            if (!File.Exists("C:\\Windows\\System\\password.txt"))
            {
                passtext = new FileStream("C:\\Windows\\System\\password.txt", FileMode.Create);
                StreamWriter checkok = new StreamWriter(passtext); checkok.Write("test2019"); checkok.Flush(); checkok.Close(); passcheck.Close();
            }

            // C:\ProgramData
            if (!File.Exists("C:\\Windows\\System32\\dcheck.txt"))
            {
                passcheck = new FileStream("C:\\Windows\\System32\\dcheck.txt", FileMode.Create);
                StreamWriter checkok = new StreamWriter(passcheck); checkok.Write("PASS"); checkok.Flush(); checkok.Close(); passcheck.Close();
            }
            else
            {
                passcheck = new FileStream("C:\\Windows\\System32\\dcheck.txt", FileMode.Open);
                StreamReader sr3 = new StreamReader(passcheck); checkp = sr3.ReadLine(); passcheck.Close();
                if (checkp.IndexOf("FAIL") != -1) { Lock pwprd = new Lock();pwprd.ShowDialog(); }

            }


            _contentDescriptions = new ObservableCollection<string>();
            ErrBox.ItemsSource = _contentDescriptions;
            _contentDescriptions2 = new ObservableCollection<string>();
            list1.ItemsSource = _contentDescriptions2;
            test.IsEnabled = false;
            //test.Focus();
            con.ConnectionString = "server=WIN-20180329ZKY\\SQLEXPRESS;database=Danfossdata;uid=sa;pwd=sqlte";
            try
            {
                con.Open();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            try
            {
                con.Close();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
            
        }





       private int flaggo = 1;
        private void DoWork_Handler(object sender, DoWorkEventArgs args)
        {
            int i;
            
            BackgroundWorker worker = sender as BackgroundWorker;
            //MessageBox.Show("worker " + dataGridView1.SelectedIndex.ToString());
            /*   flaggo = 1;
               i = 0;
               worker.ReportProgress(i);
               while (flaggo == 1)
               {
                   //  Delay(10);
                   System.Threading.Thread.Sleep(500);
               }*/
            if (single == 0)
            {
                for (i = 0; i < Item; i++)
                {
                    flaggo = 1;
                    if (worker.CancellationPending)
                    { args.Cancel = true; break; }
                    else
                    {
                        manualReset.WaitOne();//
                        worker.ReportProgress(i);
                        //    if (i == 0) { MessageBox.Show("wait"); }
                        while (flaggo == 1)
                        {
                            System.Threading.Thread.Sleep(100);
                            // MessageBox.Show("wait");
                        }

                        // System.Threading.Thread.Sleep(500);
                        // Delay(500);
                    }

                }
            }

            if (single == 1)
            {
                worker.ReportProgress(singleitem);
                
              //  worker.ReportProgress(4);
            
                /*  for (i = 0; i < dataGridView1.SelectedIndex + 3; i++)
                {
                    flaggo = 1;
                    MessageBox.Show(dataGridView1.SelectedIndex.ToString());
                    manualReset.WaitOne();
                    worker.ReportProgress(i);
                    while (flaggo == 1)
                    {
                        System.Threading.Thread.Sleep(100);
                        // MessageBox.Show("wait");
                    }
                }
               */
                // worker.ReportProgress(-1);
                //worker.ReportProgress(0);
            }

        }
        private void ProgressChanged_Handler(object sender, ProgressChangedEventArgs args)
        {
            // Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), args.ProgressPercentage.);
          //  MessageBox.Show("Change "+args.ProgressPercentage.ToString());
            StringBuilder temp = new StringBuilder(500);
            string canerrinfo, item1 = "Item",vbat,v34,distance,ccid,crid,canindex,cclength,ccdata, led = "", ledact = "", nc1, no2, nc3, no4, nc1act = "", no2act = "", nc3act = "", no4act = "";
            int passitem = 1,passitem_can=1,passitem_nc1=1, passitem_no2 = 1, passitem_nc3 = 1, passitem_no4 = 1, passitem_led = 1, passitem_dcr1=1, passitem_dcr2=1, passitem_result = 1,discount=0;
          
            uint dtime;
            byte[] sendData = null;
            double ledresult, nc1data, no2data, nc3data, no4data;
            item1 = item1 + Seq[args.ProgressPercentage].ToString();
         //   MessageBox.Show(item1);
            dataGridView1.SelectedIndex = args.ProgressPercentage;
            GetPrivateProfileString(item1, "Item", "N/A", temp, 500, ff);
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), temp.ToString());
            GetPrivateProfileString(item1, "Delay", "500", temp, 500, ff);
            dtime = uint.Parse(temp.ToString());
            GetPrivateProfileString(item1, "Vbat", "0", temp, 500, ff);
            vbat = temp.ToString();
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "Vbat=" + vbat);            
            if (vbat != "0")
            {
                sendData = Encoding.UTF8.GetBytes("VSET1:" + vbat);
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("ISET1:3");
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("OUT1");
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("VOUT1?");
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
               // Delay(dtime);
            }
            sendData = null;
            GetPrivateProfileString(item1, "V34", "0", temp, 500, ff);
            v34 = temp.ToString();
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "V34=" + v34);
            if (v34 != "0")
            {
                sendData = Encoding.UTF8.GetBytes("VSET1:" + v34);
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("ISET1:3");
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("OUT1");
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("VOUT1?");
                power1sp.Write(sendData, 0, sendData.Length);
                Delay(60);
               // Delay(dtime);
            }
            if ((vbat == "0") && (v34 == "0") && (MODEL.Content.ToString().IndexOf("CLS")==-1) )
            { errcode= outport("Dev1/port0", 0x1F); }
            if ((vbat != "0") && (v34 != "0") && (MODEL.Content.ToString().IndexOf("CLS")) == -1)
            { errcode = outport("Dev1/port0", 0x1A);  }
            if ((vbat != "0") && (v34 == "0") && (MODEL.Content.ToString().IndexOf("CLS")) == -1)
            { errcode = outport("Dev1/port0", 0x1E);  }
            if ((vbat == "0") && (v34 != "0") && (MODEL.Content.ToString().IndexOf("CLS")) == -1)
            { errcode = outport("Dev1/port0", 0x1B);  }
            if ((vbat != "0") && (v34 == "0") && (MODEL.Content.ToString().IndexOf("0118-CLS")) != -1)
            { errcode = outport("Dev3/port0", 0x1F);  }
            if ((vbat != "0") && (v34 == "0") && (MODEL.Content.ToString().IndexOf("0123-CLS")) != -1)
            { errcode = outport("Dev3/port0", 0x1A);  }
            if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcode); }
            GetPrivateProfileString(item1, "Distance", "0", temp, 500, ff);
            distance = temp.ToString();
            if (int.Parse(distance) > 0)
            {
                sp.DiscardInBuffer();
                sp.DiscardOutBuffer();
                sendcomm(sp, "ffaa030600000000b2");
                Delay(500);
                errcode = sendcomm(sp, "FFAA030401010000B2");
                if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "StepMotor Com Err:" + errcode); }
                dishex = motor_Dis(distance);
                errcode = sendcomm(sp, dishex);
                if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "StepMotor Com Err:" + errcode); }
                errcode = sendcomm(sp, "FFAA030900000000B5");
                if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "StepMotor Com Err:" + errcode); }
            }
            if (int.Parse(distance) < 0)
            {
                sp.DiscardInBuffer();
                sp.DiscardOutBuffer();
                sendcomm(sp, "ffaa030600000000b2");
                Delay(500);
                errcode = sendcomm(sp, "FFAA030400010000B1");
                if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "StepMotor Com Err:" + errcode); }
                distance = (0 - int.Parse(distance)).ToString();

                //   MessageBox.Show(distance);
                dishex = motor_Dis(distance);
                errcode = sendcomm(sp, dishex);
                if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "StepMotor Com Err:" + errcode); }
                errcode = sendcomm(sp, "FFAA030900000000B5");
                if (errcode != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "StepMotor Com Err:" + errcode); }
            }
            if (int.Parse(distance) != 0) { Delay(dtime); }
            string sendcandata = "", str1buf, recivecandata = "", crlengthbuf,lvdt = "", lvdtmax = "", lvdtmin = "", dcr1 = "", dcr2 = "", dcr1max = "", dcr1min = "", dcr2max = "", dcr2min = "";
            int tlen = 0, hexTooct = 0, ican = 0, ind, ibuf;
            Ecan.CAN_OBJ sendbusinfo = new Ecan.CAN_OBJ(), recbusinfo = new Ecan.CAN_OBJ();
            Ecan.CAN_OBJ[] recbusinfo2;
            Ecan.CAN_OBJ[] recbusinfo5= new Ecan.CAN_OBJ[1];
            GetPrivateProfileString(item1, "CCID", "", temp, 500, ff); ccid = temp.ToString();
            GetPrivateProfileString(item1, "CRID", "", temp, 500, ff); crid = temp.ToString();
            GetPrivateProfileString(item1, "Canchannl", "", temp, 500, ff); canindex = temp.ToString();
            GetPrivateProfileString(item1, "CRLENGTH", "0", temp, 500, ff); crlengthbuf = temp.ToString();
            string[] seqcrid, crinfoid;
            UInt32 canindexint;
            canindexint = UInt32.Parse(canindex);
            int cridcount = 0, px = 0;
            if (crid != "")
            {
                seqcrid = crid.ToString().Split(',');
                //int xx = 0;
                foreach (string azu in seqcrid) { cridcount++; }
                crinfoid = new string[cridcount];
                foreach (string azu in seqcrid) { crinfoid[px] = azu; px++; }
            }
            else { crinfoid = new string[cridcount + 1]; }
            recbusinfo2 = new Ecan.CAN_OBJ[cridcount + 1];
            if (ccid == "FF")
            {
                canerrinfo = canreceive2(out recbusinfo2, cridcount,canindexint);
                // MessageBox.Show(recbusinfo2[0].ID.ToString());
                // MessageBox.Show(recbusinfo2[1].ID.ToString());
                for (int pfx = 0; pfx < cridcount; pfx++)
                //  canerrinfo = cansend(sendbusinfo, out recbusinfo);
                {
                    if (canerrinfo != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), canerrinfo); }
                    if (canerrinfo == "")
                    {
                       // ErrBox.Items.Add("CanRecID :" + recbusinfo2[pfx].ID.ToString("X"));
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "CanRecID :" + recbusinfo2[pfx].ID.ToString("X"));
                    }
                }
            }
            if ((ccid != "") && (ccid != "FF"))
            {
                GetPrivateProfileString(item1, "CCLENGTH", "0", temp, 500, ff); cclength = temp.ToString();
                sendbusinfo.SendType = 0; sendbusinfo.data = new byte[8]; sendbusinfo.Reserved = new byte[3]; sendbusinfo.RemoteFlag = 0; sendbusinfo.ExternFlag = 1;
                sendbusinfo.DataLen = Convert.ToByte(int.Parse(cclength));
                GetPrivateProfileString(item1, "CCDATA", "", temp, 500, ff); ccdata = temp.ToString();
                tlen = sendbusinfo.DataLen - 1;
                for (ican = 0; ican <= tlen; ican++)
                { sendbusinfo.data[ican] = Convert.ToByte(ccdata.Substring(0 + ican * 2, 2), 0X10); sendcandata = sendcandata + sendbusinfo.data[ican].ToString("X2") + " "; }
                sendbusinfo.ID = Convert.ToUInt32(ccid, 16);
                /*  sendbusinfo.data[0] = Convert.ToByte("01", 0X10);
                  sendbusinfo.data[1] = Convert.ToByte("01", 0X10);
                  sendbusinfo.data[2] = Convert.ToByte("02", 0X10);
                  sendbusinfo.data[3] = Convert.ToByte("01", 0X10);
                  sendbusinfo.ID = Convert.ToUInt32("4012010", 16);*/
               // listBox1.Items.Add("CCID :" + sendbusinfo.ID.ToString("X")); listBox1.Items.Add("CCDATA :" + sendcandata);
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CCID :" + sendbusinfo.ID.ToString("X"));
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CCDATA :" + sendcandata);
                GetPrivateProfileString(item1, "Canmodel", "", temp, 500, ff); canmodel = temp.ToString();
                // canerrinfo = cansend(sendbusinfo, out recbusinfo);
                canerrinfo = cansend2(sendbusinfo, out recbusinfo2, cridcount, canindexint);
                for (int pfx = 0; pfx < cridcount; pfx++)
                {
                    if (canerrinfo != "") { Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), canerrinfo); }
                    if (canerrinfo == "")
                    {
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), "CanRecID :" + recbusinfo2[pfx].ID.ToString("X"));
                        // ErrBox.Items.Add("CanRecID :" + recbusinfo2[pfx].ID.ToString("X"));

                    }
                }
            }
            string actidstr;
            actidstr = "Act CRID:";
            for (int pfx = 0; pfx < cridcount; pfx++)
            { actidstr = actidstr + " " + recbusinfo2[pfx].ID.ToString("X"); }
            //     listBox1.Items.Add(actidstr);
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), actidstr);
            float canfresult = 0;
            //   recivecandata = "";
            string cistatue = "", LVDTmax = "", LVDTmin = "", LVDTstatue = "xxxx";
            string[] seq, seqx, seq2, swv;
            seqx = new string[8];
            swv = new string[3];
            int xx = 0, pp = 0, idright = 0, pf = 0, pfxx = 0;
            if (crid != "")
            {
                for (pfxx = 0; pfxx < cridcount; pfxx++)
                {
                    idright = 0;
                    pf = 0;
                    //  MessageBox.Show(cridcount.ToString());
                    do
                    {
                        //  MessageBox.Show(pf.ToString());
                        if (crinfoid[pfxx].ToUpper() == recbusinfo2[pf].ID.ToString("X").ToUpper())
                        {
                            idright = 1;
                            tlen = recbusinfo2[pf].DataLen - 1;
                            for (ican = 0; ican <= tlen; ican++)
                            { recivecandata = recivecandata + recbusinfo2[pf].data[ican].ToString("X2") + " "; }
                          //  listBox1.Items.Add("CRID :" + recbusinfo2[pf].ID.ToString("X")); listBox1.Items.Add("CRDATA :" + recivecandata);
                            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRID :" + recbusinfo2[pf].ID.ToString("X"));
                            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recivecandata);
                            GetPrivateProfileString(item1, "LVDTMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                            GetPrivateProfileString(item1, "LVDTMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                            if (((lvdtmax != "") || (lvdtmin != "")) && (MODEL.Content.ToString().IndexOf("CI") != -1))
                            {
                                GetPrivateProfileString(item1, "Canmodel", "", temp, 500, ff); canmodel = temp.ToString();
                                if (canmodel == "SW version")
                                {
                                    seq = lvdtmax.ToString().Split(',');
                                    //int xx = 0;
                                    foreach (string azu in seq) { seqx[xx] = azu; xx++; }

                                    for (pp = 0; pp < 3; pp++) { swv[pp] = seqx[pp]; }
                                    for (pp = 0; pp < 3; pp++)
                                    {
                                        if (recbusinfo2[pf].data[pp + 2].ToString("X2") != swv[pp]) { pass = 0; passitem = 0; testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                        //  dataGridView1.Rows[i].Cells[7].Value = dataGridView1.Rows[i].Cells[7].Value + recbusinfo2[pf].data[pp + 2].ToString("X2") + " "; listBox1.Items.Add("SW Verion :" + recbusinfo2[pf].data[pp + 2].ToString("X2") + " Spec :" + swv[pp]);
                                        testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString() + recbusinfo2[pf].data[pp + 2].ToString("X2") + " ";
                                        dataGridView1.ItemsSource = null;//这里是关键
                                        dataGridView1.ItemsSource = testdata;
                                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "SW Verion :" + recbusinfo2[pf].data[pp + 2].ToString("X2") + " Spec :" + swv[pp]);
                                    }

                                }
                                if (canmodel == "HW Status")
                                {
                                    Delay(dtime);
                                    if (recbusinfo2[pf].data[0].ToString("X2") != lvdtmax) { pass = 0; passitem = 0; testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                    //  dataGridView1.Rows[i].Cells[7].Value = recbusinfo2[pf].data[0].ToString("X2"); listBox1.Items.Add("HW Status :" + recbusinfo2[pf].data[0].ToString("X2"));
                                    testdata[args.ProgressPercentage].candata = recbusinfo2[pf].data[0].ToString("X2");
                                    dataGridView1.ItemsSource = null;//这里是关键
                                    dataGridView1.ItemsSource = testdata;
                                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "HW Status :" + recbusinfo2[pf].data[0].ToString("X2"));
                                }
                                if (canmodel == "D2 5V")
                                {
                                    lvdt = recbusinfo2[pf].data[7].ToString("X2") + recbusinfo2[pf].data[6].ToString("X2"); hexTooct = hextodec(lvdt);
                                    if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                 //   dataGridView1.Rows[i].Cells[7].Value = hexTooct.ToString(); listBox1.Items.Add("D2 5V :" + hexTooct.ToString());
                                    testdata[args.ProgressPercentage].candata = hexTooct.ToString();
                                    dataGridView1.ItemsSource = null;//这里是关键
                                    dataGridView1.ItemsSource = testdata;
                                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "D2 5V :" + hexTooct.ToString());
                                }

                                /*seq = lvdtmax.ToString().Split(',');

                            foreach(string azu in seq) { seqx[xx] = azu;xx++; }
                            lvdtmax = seqx[1]; LVDTstatue = seqx[0];xx = 0; seq = lvdtmin.ToString().Split(',');
                            foreach (string azu in seq) { seqx[xx] = azu; xx++; }
                            lvdtmin = seqx[1];
                            if (recbusinfo.data[2].ToString("X2") == "30") { cistatue = "blocked"; }
                            if (recbusinfo.data[2].ToString("X2") == "33") { cistatue = "floating"; }
                            if (recbusinfo.data[2].ToString("X2") == "31") { cistatue = "retracted";hexTooct= int.Parse(recbusinfo.data[0].ToString("X2"), System.Globalization.NumberStyles.HexNumber)-128; }
                            if (recbusinfo.data[2].ToString("X2") == "32") { cistatue = "extended"; hexTooct = -(int.Parse(recbusinfo.data[1].ToString("X2"), System.Globalization.NumberStyles.HexNumber) - 128); }
                            if (cistatue!=LVDTstatue) { pass = 0; passitem = 0; }
                            dataGridView1.Rows[i].Cells[7].Value = cistatue; listBox1.Items.Add("LVDTstatue :" + cistatue);
                            if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; }
                            dataGridView1.Rows[i].Cells[7].Value = hexTooct.ToString(); listBox1.Items.Add("LVDT :" + hexTooct.ToString());
*/
                            }
                            int poserr = 1;
                            discount = 0;
                            if (((lvdtmax != "") || (lvdtmin != "")) && (MODEL.Content.ToString().IndexOf("CI") == -1))
                            {
                                if (MODEL.Content.ToString().IndexOf("CLS") == -1)
                                {
                                    lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                    if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                    // dataGridView1.Rows[i].Cells[7].Value = hexTooct.ToString(); listBox1.Items.Add("LVDT :" + hexTooct.ToString());
                                   
                                }
                                if ((MODEL.Content.ToString().IndexOf("CLS") != -1)&&(canmodel.IndexOf("Middle")!=-1))
                                {
                                    lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                 //   MessageBox.Show(hexTooct.ToString());
                                    if ((hexTooct > int.Parse(lvdtmax)+500) || (hexTooct < int.Parse(lvdtmin)-500)) { pass = 0; passitem = 0; passitem_can = 0; }
                                    else
                                    {
                                        do
                                        {
                                          //  MessageBox.Show(hexTooct.ToString());
                                            sp.DiscardInBuffer();
                                            sp.DiscardOutBuffer();
                                            if (hexTooct > int.Parse(lvdtmax))
                                            {
                                                sendcomm(sp, "ffaa030600000000b2");
                                                Delay(500);
                                                errcode = sendcomm(sp, "FFAA030400010000B1");
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                              //  distance = discount.ToString();
                                                dishex = motor_Dis("1");
                                                errcode = sendcomm(sp, dishex);
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                                errcode = sendcomm(sp, "FFAA030900000000B5");
                                              //  MessageBox.Show("DIS: " + distance);
                                                Delay(2000);
                                                cansend2(sendbusinfo, out recbusinfo5, cridcount, canindexint);
                                                lvdt = recbusinfo5[0].data[1].ToString("X2") + recbusinfo5[0].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recbusinfo5[0].data[0].ToString("X2") +" "+ recbusinfo5[0].data[1].ToString("X2"));
                                                //   pass = 0; passitem = 0; passitem_can = 0;
                                                // MessageBox.Show(hexTooct.ToString());
                                            }
                                            if (hexTooct < int.Parse(lvdtmin))
                                            {
                                                sendcomm(sp, "ffaa030600000000b2");
                                                Delay(500);
                                                errcode = sendcomm(sp, "FFAA030401010000B2");
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                             //   distance = discount.ToString();
                                                dishex = motor_Dis("1");
                                                errcode = sendcomm(sp, dishex);
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                                errcode = sendcomm(sp, "FFAA030900000000B5");
                                              //  MessageBox.Show("DIS: " + distance);
                                                Delay(2000);
                                                cansend2(sendbusinfo, out recbusinfo5, cridcount, canindexint);
                                                lvdt = recbusinfo5[0].data[1].ToString("X2") + recbusinfo5[0].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recbusinfo5[0].data[0].ToString("X2") + " " + recbusinfo5[0].data[1].ToString("X2"));
                                                //   pass = 0; passitem = 0; passitem_can = 0;
                                                //MessageBox.Show(hexTooct.ToString());
                                            }
if((hexTooct > int.Parse(lvdtmin))&&(hexTooct < int.Parse(lvdtmax)))
                                            {
                                                poserr = 0; //pass = 1; passitem = 1; passitem_can = 1;
                                            }
                                            discount++;
                                        } while ((discount < 10) && (poserr == 1));


                                    }//else

                                }//Middle


                                if ((MODEL.Content.ToString().IndexOf("CLS") != -1) && (canmodel.IndexOf("Potting") != -1))
                                {
                                    lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                 //   MessageBox.Show(hexTooct.ToString());
                                    if ((hexTooct > int.Parse(lvdtmax) + 500) || (hexTooct < int.Parse(lvdtmin) - 500)) { pass = 0; passitem = 0; passitem_can = 0; }
                                    else
                                    {
                                        do
                                        {
                                            sp.DiscardInBuffer();
                                            sp.DiscardOutBuffer();
                                            if (hexTooct > int.Parse(lvdtmax))
                                            {
                                                sendcomm(sp, "ffaa030600000000b2");
                                                Delay(500);
                                                errcode = sendcomm(sp, "FFAA030400010000B1");
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                              //  distance = discount.ToString();
                                                dishex = motor_Dis("1");
                                                errcode = sendcomm(sp, dishex);
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                                errcode = sendcomm(sp, "FFAA030900000000B5");
                                              //  MessageBox.Show("DIS: " + distance);
                                                Delay(1000);
                                                cansend2(sendbusinfo, out recbusinfo5, cridcount, canindexint);
                                                lvdt = recbusinfo5[0].data[1].ToString("X2") + recbusinfo5[0].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recbusinfo5[0].data[0].ToString("X2") + " " + recbusinfo5[0].data[1].ToString("X2"));
                                                //  pass = 0; passitem = 0; passitem_can = 0;
                                                //  MessageBox.Show(hexTooct.ToString());
                                            }
                                            if (hexTooct < int.Parse(lvdtmin))
                                            {
                                                sendcomm(sp, "ffaa030600000000b2");
                                                Delay(500);
                                                errcode = sendcomm(sp, "FFAA030401010000B2");
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                             //   distance = discount.ToString();
                                                dishex = motor_Dis("1");
                                                errcode = sendcomm(sp, dishex);
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                                errcode = sendcomm(sp, "FFAA030900000000B5");
                                             //   MessageBox.Show("DIS: " + distance);
                                                Delay(1000);
                                                cansend2(sendbusinfo, out recbusinfo5, cridcount, canindexint);
                                                lvdt = recbusinfo5[0].data[1].ToString("X2") + recbusinfo5[0].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recbusinfo5[0].data[0].ToString("X2") + " " + recbusinfo5[0].data[1].ToString("X2"));
                                                // pass = 0; passitem = 0; passitem_can = 0;
                                                // MessageBox.Show(hexTooct.ToString());
                                            }
                                            if ((hexTooct > int.Parse(lvdtmin)) && (hexTooct < int.Parse(lvdtmax)))
                                            {
                                                poserr = 0; //pass = 1; passitem = 1; passitem_can = 1;
                                            }
                                            discount++;
                                        } while ((discount < 10) && (poserr == 1));


                                    }//else

                                }//Potting

                                if ((MODEL.Content.ToString().IndexOf("CLS") != -1) && (canmodel.IndexOf("LED") != -1))
                                {
                                    lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                   // MessageBox.Show(hexTooct.ToString());
                                    if ((hexTooct > int.Parse(lvdtmax) + 500) || (hexTooct < int.Parse(lvdtmin) - 500)) { pass = 0; passitem = 0; passitem_can = 0; }
                                    else
                                    {
                                        do
                                        {
                                            sp.DiscardInBuffer();
                                            sp.DiscardOutBuffer();
                                            if (hexTooct > int.Parse(lvdtmax))
                                            {
                                                sendcomm(sp, "ffaa030600000000b2");
                                                Delay(500);
                                                errcode = sendcomm(sp, "FFAA030400010000B1");
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                              //  distance = discount.ToString();
                                                dishex = motor_Dis("1");
                                                errcode = sendcomm(sp, dishex);
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                                errcode = sendcomm(sp, "FFAA030900000000B5");
                                               // MessageBox.Show("DIS: "+ distance);
                                                Delay(1000);
                                                cansend2(sendbusinfo, out recbusinfo5, cridcount, canindexint);
                                                lvdt = recbusinfo5[0].data[1].ToString("X2") + recbusinfo5[0].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recbusinfo5[0].data[0].ToString("X2") + " " + recbusinfo5[0].data[1].ToString("X2"));
                                                //    pass = 0; passitem = 0; passitem_can = 0;
                                                //   MessageBox.Show(hexTooct.ToString());
                                            }
                                            if (hexTooct < int.Parse(lvdtmin))
                                            {
                                                sendcomm(sp, "ffaa030600000000b2");
                                                Delay(500);
                                                errcode = sendcomm(sp, "FFAA030401010000B2");
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                               // distance = (int.Parse(distance) + 1).ToString();
                                                dishex = motor_Dis("1");                                              
                                                errcode = sendcomm(sp, dishex);
                                                Delay(1000);
                                                sp.DiscardInBuffer();
                                                sp.DiscardOutBuffer();
                                                errcode = sendcomm(sp, "FFAA030900000000B5");
                                             //   MessageBox.Show("DIS: " + distance);
                                                Delay(1000);
                                                cansend2(sendbusinfo, out recbusinfo5, cridcount, canindexint);
                                                lvdt = recbusinfo5[0].data[1].ToString("X2") + recbusinfo5[0].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRDATA :" + recbusinfo5[0].data[0].ToString("X2") + " " + recbusinfo5[0].data[1].ToString("X2"));
                                                // pass = 0; passitem = 0; passitem_can = 0;
                                                //  MessageBox.Show(hexTooct.ToString());
                                            }
                                            if ((hexTooct > int.Parse(lvdtmin)) && (hexTooct < int.Parse(lvdtmax)))
                                            {
                                                poserr = 0; //pass = 1; passitem = 1; passitem_can = 1;
                                            }
                                            discount++;
                                        } while ((discount < 10) && (poserr == 1));


                                    }//else

                                }//LED


                                sendcomm(sp, "ffaa030600000000b2");
                                Delay(500);
                                if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                testdata[args.ProgressPercentage].candata = hexTooct.ToString();
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LVDT :" + hexTooct.ToString());
                            }//lvdt pos
                            for (ibuf = 1; ibuf < 9; ibuf++)
                            {
                                str1buf = "AD" + ibuf.ToString() + "MAX"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmax = temp.ToString();
                                str1buf = "AD" + ibuf.ToString() + "MIN"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmin = temp.ToString();
                                if ((lvdtmax != "") || (lvdtmin != ""))
                                {
                                    if (ibuf % 2 != 0) { lvdt = recbusinfo2[pf].data[ibuf].ToString("X2") + recbusinfo2[pf].data[ibuf - 1].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f / 1000; }
                                    if (ibuf % 2 == 0) { lvdt = recbusinfo2[pf].data[ibuf - 1].ToString("X2") + recbusinfo2[pf].data[ibuf - 2].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f / 1000; }
                                    if ((canfresult > float.Parse(lvdtmax)) || (canfresult < float.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can=0; }
                                   // dataGridView1.Rows[i].Cells[8 + ibuf - 1].Value = canfresult.ToString(); str1buf = "AD" + ibuf.ToString() + " :"; listBox1.Items.Add(str1buf + canfresult.ToString());
                                    testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString() + canfresult.ToString() + ",";
                                    dataGridView1.ItemsSource = null;//这里是关键
                                    dataGridView1.ItemsSource = testdata;
                                    if (passitem_can != 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                    if (passitem_can == 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                    str1buf = "AD" + ibuf.ToString() + " :";
                                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), str1buf + canfresult.ToString());
                                }
                            }
                            for (ibuf = 1; ibuf < 5; ibuf++)
                            {
                                str1buf = "ILS" + ibuf.ToString() + "MAX"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmax = temp.ToString();
                                str1buf = "ILS" + ibuf.ToString() + "MIN"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmin = temp.ToString();
                                if ((lvdtmax != "") || (lvdtmin != ""))
                                {
                                    if (ibuf % 2 != 0) { lvdt = recbusinfo2[pf].data[ibuf].ToString("X2") + recbusinfo2[pf].data[ibuf - 1].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f; }
                                    if (ibuf % 2 == 0) { lvdt = recbusinfo2[pf].data[ibuf - 1].ToString("X2") + recbusinfo2[pf].data[ibuf - 2].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f; }
                                    if ((canfresult-5000 > float.Parse(lvdtmax)) || (canfresult-5000 < float.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can=0; }
                                    //dataGridView1.Rows[i].Cells[16 + ibuf - 1].Value = canfresult.ToString(); str1buf = "ILS" + ibuf.ToString() + " :"; listBox1.Items.Add(str1buf + canfresult.ToString());
                                    testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString() + (canfresult-5000).ToString() + ",";
                                    dataGridView1.ItemsSource = null;//这里是关键
                                    dataGridView1.ItemsSource = testdata;
                                    if (passitem_can != 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                    if (passitem_can == 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                    str1buf = "ILS" + ibuf.ToString() + " :";
                                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), str1buf + (canfresult-5000).ToString());
                                }
                            }

                            GetPrivateProfileString(item1, "XPMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                            GetPrivateProfileString(item1, "XPMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                            if ((lvdtmax != "") || (lvdtmin != ""))
                            {
                                lvdt = recbusinfo2[pf].data[3].ToString("X2") + recbusinfo2[pf].data[2].ToString("X2"); x_p = hextodec(lvdt);
                                if ((x_p > int.Parse(lvdtmax)) || (x_p < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                testdata[args.ProgressPercentage].candata = x_p.ToString();
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "X_P :" + x_p.ToString());
                            }
                            for (ibuf = 1; ibuf < 5; ibuf++)
                            {
                                if (ibuf == 1)
                                {
                                    GetPrivateProfileString(item1, "NC1CurrMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                                    GetPrivateProfileString(item1, "NC1CurrMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                                    if ((lvdtmax != "") || (lvdtmin != ""))
                                    {
                                        //  MessageBox.Show(lvdt);
                                        lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                        // MessageBox.Show(hexTooct.ToString());
                                        if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                        testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString()+ hexTooct.ToString()+",";
                                        dataGridView1.ItemsSource = null;//这里是关键
                                        dataGridView1.ItemsSource = testdata;
                                        if (passitem_can != 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                        if (passitem_can == 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NC1Curr :" + hexTooct.ToString());
                                    }
                                }
                               if(ibuf==2)
                                {
                                    GetPrivateProfileString(item1, "NO2CurrMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                                    GetPrivateProfileString(item1, "NO2CurrMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                                    if ((lvdtmax != "") || (lvdtmin != ""))
                                    {
                                        lvdt = recbusinfo2[pf].data[3].ToString("X2") + recbusinfo2[pf].data[2].ToString("X2"); hexTooct = hextodec(lvdt);
                                        if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                        testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString()+ hexTooct.ToString()+",";
                                        dataGridView1.ItemsSource = null;//这里是关键
                                        dataGridView1.ItemsSource = testdata;
                                        if (passitem_can != 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                        if (passitem_can == 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NO2Curr :" + hexTooct.ToString());
                                    }

                                }
                                if (ibuf == 3)
                                {
                                    GetPrivateProfileString(item1, "NC3CurrMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                                    GetPrivateProfileString(item1, "NC3CurrMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                                    if ((lvdtmax != "") || (lvdtmin != ""))
                                    {
                                        lvdt = recbusinfo2[pf].data[5].ToString("X2") + recbusinfo2[pf].data[4].ToString("X2"); hexTooct = hextodec(lvdt);
                                        if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                        testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString()+hexTooct.ToString()+",";
                                        dataGridView1.ItemsSource = null;//这里是关键
                                        dataGridView1.ItemsSource = testdata;
                                        if (passitem_can != 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                        if (passitem_can == 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NC3Curr :" + hexTooct.ToString());
                                    }

                                }
                                if (ibuf == 4)
                                {
                                    GetPrivateProfileString(item1, "NO4CurrMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                                    GetPrivateProfileString(item1, "NO4CurrMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                                    if ((lvdtmax != "") || (lvdtmin != ""))
                                    {
                                        lvdt = recbusinfo2[pf].data[7].ToString("X2") + recbusinfo2[pf].data[6].ToString("X2"); hexTooct = hextodec(lvdt);
                                        if ((hexTooct > int.Parse(lvdtmax)) || (hexTooct < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                        testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString()+ hexTooct.ToString()+",";
                                        dataGridView1.ItemsSource = null;//这里是关键
                                        dataGridView1.ItemsSource = testdata;
                                        if (passitem_can != 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                        if (passitem_can == 0)
                                        { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NO4Curr :" + hexTooct.ToString());
                                    }

                                }


                            }

                          

                            

                           




                            GetPrivateProfileString(item1, "XSMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                            GetPrivateProfileString(item1, "XSMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                            if ((lvdtmax != "") || (lvdtmin != ""))
                            {
                                lvdt = recbusinfo2[pf].data[3].ToString("X2") + recbusinfo2[pf].data[2].ToString("X2"); x_s = hextodec(lvdt);
                                if ((x_s > int.Parse(lvdtmax)) || (x_s < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can = 0; }
                                testdata[args.ProgressPercentage].candata = x_s.ToString();
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "X_S :" + x_s.ToString());
                            }

                            GetPrivateProfileString(item1, "Canmodel", "", temp, 500, ff); canmodel = temp.ToString();
                            if(canmodel=="Temp_OffSetP")
                            {
                             //   MessageBox.Show(x_p.ToString());
                                lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                if (hexTooct == 65535) { pass = 0; passitem = 0; passitem_can = 0; }
                              //  MessageBox.Show(passitem_can.ToString());
                                testdata[args.ProgressPercentage].candata = hexTooct.ToString();
                                lvdt = recbusinfo2[pf].data[3].ToString("X2") + recbusinfo2[pf].data[2].ToString("X2"); hexTooct = hextodec(lvdt);
                                if ((hexTooct > (x_p+20)) || (hexTooct  <(x_p-20))) { pass = 0; passitem = 0; passitem_can = 0; }
                            //    MessageBox.Show(passitem_can.ToString());
                                testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata+","+ hexTooct;
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "DTSOffset&DTSTemp_P :" + testdata[args.ProgressPercentage].candata);
                            }
                            if (canmodel == "Temp_OffSetS")
                            {
                                lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt);
                                if (hexTooct == 65535) { pass = 0; passitem = 0; passitem_can = 0; }
                                testdata[args.ProgressPercentage].candata = hexTooct.ToString();
                                lvdt = recbusinfo2[pf].data[3].ToString("X2") + recbusinfo2[pf].data[2].ToString("X2"); hexTooct = hextodec(lvdt);
                                if ((hexTooct > (x_s + 20)) || (hexTooct < (x_s - 20))) { pass = 0; passitem = 0; passitem_can = 0; }
                                testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata + "," + hexTooct;
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "DTSOffset&DTSTemp_S :" + testdata[args.ProgressPercentage].candata);
                            }

                            GetPrivateProfileString(item1, "IHS1MAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                            GetPrivateProfileString(item1, "IHS1MIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                            if ((lvdtmax != "") || (lvdtmin != ""))
                            {
                                lvdt = recbusinfo2[pf].data[5].ToString("X2") + recbusinfo2[pf].data[4].ToString("X2"); hexTooct = hextodec(lvdt);
                                if ((hexTooct-5000 > int.Parse(lvdtmax)) || (hexTooct-5000 < int.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can=0; }
                              //  dataGridView1.Rows[i].Cells[20].Value = hexTooct.ToString(); listBox1.Items.Add("ISH1 :" + hexTooct.ToString());
                                testdata[args.ProgressPercentage].candata= (hexTooct-5000).ToString();
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "ISH1 :" + (hexTooct-5000).ToString());
                            }
                            for (ibuf = 1; ibuf < 3; ibuf++)
                            {
                                str1buf = "FIN" + ibuf.ToString() + "DCMAX"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmax = temp.ToString();
                                str1buf = "FIN" + ibuf.ToString() + "DCMIN"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmin = temp.ToString();
                                if ((lvdtmax != "") || (lvdtmin != ""))
                                {
                                    if (ibuf == 1) { lvdt = recbusinfo2[pf].data[1].ToString("X2") + recbusinfo2[pf].data[0].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f / 100; }
                                    if (ibuf == 2) { lvdt = recbusinfo2[pf].data[5].ToString("X2") + recbusinfo2[pf].data[4].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f / 100; }
                                    if ((canfresult > float.Parse(lvdtmax)) || (canfresult < float.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can=0; }
                                  //  dataGridView1.Rows[i].Cells[21 + ibuf - 1].Value = canfresult.ToString(); str1buf = "FIN" + ibuf.ToString() + "DC :"; listBox1.Items.Add(str1buf + canfresult.ToString());
                                    testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString() + canfresult.ToString() + ",";
                                    dataGridView1.ItemsSource = null;//这里是关键
                                    dataGridView1.ItemsSource = testdata;
                                    if (passitem_can != 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                    if (passitem_can == 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                    str1buf = "FIN" + ibuf.ToString() + "DC :";
                                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), str1buf + canfresult.ToString());
                                }
                            }
                            for (ibuf = 1; ibuf < 3; ibuf++)
                            {
                                str1buf = "FIN" + ibuf.ToString() + "FREQMAX"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmax = temp.ToString();
                                str1buf = "FIN" + ibuf.ToString() + "FREQMIN"; GetPrivateProfileString(item1, str1buf, "", temp, 500, ff); lvdtmin = temp.ToString();
                                if ((lvdtmax != "") || (lvdtmin != ""))
                                {
                                    if (ibuf == 1) { lvdt = recbusinfo2[pf].data[3].ToString("X2") + recbusinfo2[pf].data[2].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f; }
                                    if (ibuf == 2) { lvdt = recbusinfo2[pf].data[7].ToString("X2") + recbusinfo2[pf].data[6].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f; }
                                    if ((canfresult > float.Parse(lvdtmax)) || (canfresult < float.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can=0; }
                                  //  dataGridView1.Rows[i].Cells[23 + ibuf - 1].Value = canfresult.ToString(); str1buf = "FIN" + ibuf.ToString() + "FREQ :"; listBox1.Items.Add(str1buf + canfresult.ToString());
                                    testdata[args.ProgressPercentage].candata = testdata[args.ProgressPercentage].candata.ToString() + canfresult.ToString() + ",";
                                    dataGridView1.ItemsSource = null;//这里是关键
                                    dataGridView1.ItemsSource = testdata;
                                    if (passitem_can != 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                    if (passitem_can == 0)
                                    { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                    str1buf = "FIN" + ibuf.ToString() + "FREQ :";
                                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), str1buf + canfresult.ToString());
                                }
                            }
                            GetPrivateProfileString(item1, "KEYMAX", "", temp, 500, ff); lvdtmax = temp.ToString();
                            GetPrivateProfileString(item1, "KEYMIN", "", temp, 500, ff); lvdtmin = temp.ToString();
                            if ((lvdtmax != "") || (lvdtmin != ""))
                            {
                                lvdt = recbusinfo2[pf].data[5].ToString("X2") + recbusinfo2[pf].data[4].ToString("X2"); hexTooct = hextodec(lvdt); canfresult = hexTooct * 1.0f / 1000;
                                if ((canfresult > float.Parse(lvdtmax)) || (canfresult < float.Parse(lvdtmin))) { pass = 0; passitem = 0; passitem_can=0; }
                                //dataGridView1.Rows[i].Cells[25].Value = canfresult.ToString(); listBox1.Items.Add("KEYVOL :" + canfresult.ToString());
                                testdata[args.ProgressPercentage].candata = canfresult.ToString();
                                dataGridView1.ItemsSource = null;//这里是关键
                                dataGridView1.ItemsSource = testdata;
                                if (passitem_can != 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
                                if (passitem_can == 0)
                                { testdata[args.ProgressPercentage].candataFColor = Brushes.Red; }
                                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "KEYVOL :" + canfresult.ToString());
                            }

                        }//crid.ToUpper()== recbusinfo.ID.ToString("X").ToUpper()
                         //   MessageBox.Show(pf.ToString());

                        pf++;
                    } while (pf < cridcount);
                    if (idright == 0)
                    {

                      //  listBox1.Items.Add("CRID : " + crinfoid[pfxx].ToUpper() + "don't exit");
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "CRID : " + crinfoid[pfxx].ToUpper() + "don't exit");
                        pass = 0; passitem = 0;// dataGridView1.Rows[i].Cells[7].Value = "ID ERR";
                        testdata[args.ProgressPercentage].candataFColor = Brushes.Red;
                        if (crlengthbuf != "0") { testdata[args.ProgressPercentage].candata = "ID ERR"; }
                        if (crlengthbuf == "0") { testdata[args.ProgressPercentage].candata = "None"; }
                        dataGridView1.ItemsSource = null;//这里是关键
                        dataGridView1.ItemsSource = testdata;
                    }
                }

            }//crid!=""
            if (crid == "") { testdata[args.ProgressPercentage].candata = ""; testdata[args.ProgressPercentage].candataFColor = Brushes.Green; }
            //  MessageBox.Show("22");
            if (CanEnable == "1") { Ecan.ResetCAN(4, 0, 0); Ecan.ResetCAN(4, 0, 1); }
            double dcrr;
            int cmi, ind2;
            GetPrivateProfileString(item1, "DCR1MAX", "", temp, 500, ff); dcr1max = temp.ToString();
            GetPrivateProfileString(item1, "DCR1MIN", "", temp, 500, ff); dcr1min = temp.ToString();
            if ((dcr1max != "") || (dcr1min != ""))
            {
                if ((vbat != "0") && (v34 != "0")) { outport("Dev1/port0", 0x12); }
                if ((vbat != "0") && (v34 == "0")) { outport("Dev1/port0", 0x16); }
                if ((vbat == "0") && (v34 == "0")) { outport("Dev1/port0", 0x17); }
                if ((vbat == "0") && (v34 != "0")) { outport("Dev1/port0", 0x13); }
                Delay(dtime); dcr1 = measres(DCRADD); ind = dcr1.IndexOf("E"); ind2 = dcr1.IndexOf("\n");
                //MessageBox.Show(dcr1.Length.ToString());
             //    MessageBox.Show(dcr1.Substring(1, ind-1)); MessageBox.Show(dcr1.Substring(ind + 1, ind2 - ind - 1));
                cmi = int.Parse(dcr1.Substring(ind + 1, ind2 - ind - 1));

                dcrr = double.Parse(dcr1.Substring(1, ind - 1)) * System.Math.Pow(10, cmi);
                if ((dcrr - 0.8 > float.Parse(dcr1max)) || (dcrr - 0.8 < float.Parse(dcr1min)))
                { pass = 0; passitem = 0; passitem_dcr1=0; }
                testdata[args.ProgressPercentage].dcr1 = (dcrr - 0.8).ToString("#0.000"); //listBox1.Items.Add("DCR1 :" + dcr1);
                dataGridView1.ItemsSource = null;//这里是关键
                dataGridView1.ItemsSource = testdata;
                if (passitem_dcr1 != 0)
                { testdata[args.ProgressPercentage].dcr1FColor = Brushes.Green; }
                if (passitem_dcr1 == 0)
                { testdata[args.ProgressPercentage].dcr1FColor = Brushes.Red; }
                // DataGridRow
                //  dataGridView1.DataContext = testdata;
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "DCR1 :" + dcr1);
            }
            GetPrivateProfileString(item1, "DCR2MAX", "", temp, 500, ff); dcr2max = temp.ToString();
            GetPrivateProfileString(item1, "DCR2MIN", "", temp, 500, ff); dcr2min = temp.ToString();
            if ((dcr2max != "") || (dcr2min != ""))
            {
                if ((vbat != "0") && (v34 != "0")) { outport("Dev1/port0", 0x0A); }
                if ((vbat != "0") && (v34 == "0")) { outport("Dev1/port0", 0x0E); }
                if ((vbat == "0") && (v34 == "0")) { outport("Dev1/port0", 0x0F); }
                if ((vbat == "0") && (v34 != "0")) { outport("Dev1/port0", 0x0B); }
                Delay(dtime); dcr2 = measres(DCRADD); ind = dcr2.IndexOf("E"); ind2 = dcr2.IndexOf("\n");
                cmi = int.Parse(dcr2.Substring(ind + 1, ind2 - ind - 1));
                dcrr = double.Parse(dcr2.Substring(1, ind - 1)) * System.Math.Pow(10, cmi);
                if ((dcrr - 0.8 > float.Parse(dcr2max)) || (dcrr - 0.8 < float.Parse(dcr2min)))
                { pass = 0; passitem = 0; passitem_dcr2=0; }
                //dataGridView1.Rows[i].Cells[27].Value = (dcrr - 0.8).ToString("#0.000"); listBox1.Items.Add("DCR2 :" + dcr2);
                testdata[args.ProgressPercentage].dcr2 = (dcrr - 0.8).ToString("#0.000"); //listBox1.Items.Add("DCR1 :" + dcr1);
                dataGridView1.ItemsSource = null;//这里是关键
                dataGridView1.ItemsSource = testdata;
                if (passitem_dcr2 != 0)
                { testdata[args.ProgressPercentage].dcr2FColor = Brushes.Green; }
                if (passitem_dcr2 == 0)
                { testdata[args.ProgressPercentage].dcr2FColor = Brushes.Red; }
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "DCR2 :" + dcr2);
            }

            GetPrivateProfileString(item1, "LED", "", temp, 500, ff);
            led = temp.ToString();
            double[] ledresult2;
            int ledred = 0, ledgreen = 0, ledblue = 0, ledtestcount = 0,ledwhite=0,noled=0;
            if (led != "")
            {
                Delay(dtime);
                ledresult2 = new double[1000];
                do
                {
                    ledred = 0; ledgreen = 0; ledblue = 0;
                    ledresult2 = measvol2(LEDadd, 1000);
                 //   MessageBox.Show(ledresult2.ToString());
                    for (int ledcounti = 0; ledcounti < 1000; ledcounti++)
                    {
                       // MessageBox.Show(ledresult2[ledcounti].ToString());
                        if (ledresult2[ledcounti] < 0.7) { noled++; }
                        if ((ledresult2[ledcounti] >= 1.25) && (ledresult2[ledcounti] <= 1.8)) { ledgreen++; }
                        if ((ledresult2[ledcounti] >= 2.2) && (ledresult2[ledcounti] <= 3.3)) { ledred++; }
                        if ((ledresult2[ledcounti] >3.3) && (ledresult2[ledcounti] <= 4)) { ledwhite++; }
                        if ((ledresult2[ledcounti] >= 0.7) && (ledresult2[ledcounti] <= 1.2)) { ledblue++; }
                    }
                    //   MessageBox.Show("Green :" + ledgreen.ToString());
                    ledtestcount++;
                    Delay(5);
                } while ((ledtestcount < 50) && (ledgreen < 100) && (ledred < 100) && (ledblue < 100)&&(ledwhite<100));
            //    MessageBox.Show("Green :" + ledgreen.ToString());
              //  MessageBox.Show("Red :" + ledred.ToString());
                //MessageBox.Show("No :" + noled.ToString());
                //MessageBox.Show("White :" + ledwhite.ToString());
                if ((ledgreen > 100) && (ledred < 10) && (ledblue < 10) && (ledwhite < 10))
                {
                    ledact = "Green"; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED :" + ledact + " "); testdata[args.ProgressPercentage].led=ledact; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                }
                if ((ledred > 100) && (ledgreen < 10) && (ledblue < 10) && (ledwhite < 10) )
                {
                    ledact = "Red"; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED :" + ledact + " "); testdata[args.ProgressPercentage].led = ledact; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                }
                if ((ledblue > 100) && (ledgreen < 10) && (ledred < 10) && (ledwhite < 10) )
                {
                    ledact = "Blue"; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED :" + ledact + " "); testdata[args.ProgressPercentage].led = ledact; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                }
                if ((ledwhite > 100) && (ledgreen < 10) && (ledred < 10) && (ledblue < 10) )
                {
                    ledact = "White"; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED :" + ledact + " "); testdata[args.ProgressPercentage].led = ledact; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                }
                if ((noled > 100) && (ledgreen < 10) && (ledred < 10) && (ledblue < 10) && (ledwhite < 10))
                {
                    ledact = "Nolight"; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED :" + ledact + " "); testdata[args.ProgressPercentage].led = ledact; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                }
                if (led != ledact) { pass = 0; passitem = 0;passitem_led = 0; }
                if ((ledact != "Green") && (ledact != "Red") && (ledact != "Blue")&&(ledact != "White")&&(ledact != "Nolight"))
                {
                    testdata[args.ProgressPercentage].led = "Err"; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED :Err");
                }
                if (passitem_led == 0)
                { testdata[args.ProgressPercentage].ledFColor = Brushes.Red; }
                if (passitem_led != 0)
                { testdata[args.ProgressPercentage].ledFColor = Brushes.Green; }
                /*   ledresult = measvol(LEDadd);
                if (ledresult == -1000.0) { ErrBox.Items.Add("Meas_LED_Vol err"); pass = 0; dataGridView1.Rows[i].Cells[6].Value = ledact; listBox1.Items.Add("LED :Err");passitem = 0; }                           
                if ((ledresult >= 1.25) && (ledresult <= 1.8)) { ledact = "Green"; listBox1.Items.Add("LED :" + ledact+" " + ledresult.ToString()); dataGridView1.Rows[i].Cells[6].Value = ledact; }
                if ((ledresult >= 2.48) && (ledresult <= 3.5)) { ledact = "Red"; listBox1.Items.Add("LED :" + ledact + " " + ledresult.ToString()); dataGridView1.Rows[i].Cells[6].Value = ledact; }
                if ((ledresult >= 0.7) && (ledresult <= 1.2)) { ledact = "Blue"; listBox1.Items.Add("LED :" + ledact + " " + ledresult.ToString()); dataGridView1.Rows[i].Cells[6].Value = ledact; }
                if (ledresult >3.5) { ErrBox.Items.Add("LED_Vol>3.5"); pass = 0; dataGridView1.Rows[i].Cells[6].Value = ledact; listBox1.Items.Add("LED :Err "+ ledresult.ToString()); passitem = 0; }
                if ((ledresult <0.7)&&(ledresult>-1000)) { ErrBox.Items.Add("LED_Vol<0.7"); pass = 0; dataGridView1.Rows[i].Cells[6].Value = "<0.7"; listBox1.Items.Add("LED :Err "+ ledresult.ToString()); passitem = 0; }
                if (led != ledact) { pass = 0;  passitem = 0; listBox1.Items.Add("LED Vol :" + ledresult.ToString()); }
                if((ledact!="Green")&& (ledact != "Red")&& (ledact != "Blue")) { dataGridView1.Rows[i].Cells[6].Value = "Err"; listBox1.Items.Add("LED :Err" ); }
                */
                /*  if (passitem == 0)
                  {

                      if (!Window.GetWindow(dataGridView1).IsVisible)
                      {
                          Window.GetWindow(dataGridView1).Show();
                      }

                      //   DataRowView drv = dataGridView1.Items[args.ProgressPercentage] as DataRowView;
                      DataGridRow row = (DataGridRow)this.dataGridView1.ItemContainerGenerator.ContainerFromIndex(1);
                      if (row == null)
                      {
                          MessageBox.Show("row null");
                          dataGridView1.UpdateLayout();
                          dataGridView1.ScrollIntoView(dataGridView1.Items[1]);
                          row = (DataGridRow)dataGridView1.ItemContainerGenerator.ContainerFromIndex(1);
                      }
                      DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(row);
                      DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(6);
                      if (cell != null)
                      {
                        //  MessageBox.Show("cell null");
                          dataGridView1.ScrollIntoView(row, dataGridView1.Columns[6]);
                          cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(6);
                          MessageBox.Show("cell brush");
                          cell.Background = new SolidColorBrush(Colors.Red);
                      }

                  }
  */


            }
            if (led == "")
            {
                testdata[args.ProgressPercentage].led = "None"; testdata[args.ProgressPercentage].ledFColor = Brushes.Green; dataGridView1.ItemsSource = null;//这里是关键
                dataGridView1.ItemsSource = testdata; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "LED don't care");
            }

            if (CoilNumber != "0")
            {
                GetPrivateProfileString(item1, "NC1", "", temp, 500, ff); nc1 = temp.ToString();
                GetPrivateProfileString(item1, "NO2", "", temp, 500, ff); no2 = temp.ToString();
                GetPrivateProfileString(item1, "NC3", "", temp, 500, ff); nc3 = temp.ToString();
                GetPrivateProfileString(item1, "NO4", "", temp, 500, ff); no4 = temp.ToString();

                if (nc1 != "")
                {
                   
                        Delay(dtime);
                        nc1data = measvol(NC1add);
                        // listBox1.Items.Add("NC1 :" + nc1data.ToString("#0.00"));
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NC1 :" + nc1data.ToString("#0.00"));
                        if (nc1data > 4.5) { nc1act = "OFF"; }
                        if ((nc1data < 1) && (nc1data >= 0)) { nc1act = "ON"; }
                        if ((nc1data >= 1) && (nc1data <= 4.5)) { nc1act = "ERR"; }
                        if (nc1data < 0) { nc1act = "ERR"; }
                        testdata[args.ProgressPercentage].nc1 = nc1act;
                        dataGridView1.ItemsSource = null;//这里是关键
                        dataGridView1.ItemsSource = testdata;
                        if (nc1 != nc1act) { pass = 0; passitem = 0; passitem_nc1 = 0; }
                        if (passitem_nc1 != 0)
                        { testdata[args.ProgressPercentage].nc1FColor = Brushes.Green; }
                        if (passitem_nc1 == 0)
                        { testdata[args.ProgressPercentage].nc1FColor = Brushes.Red; }
                 
                   

                    }
                if (nc1 == "")
                {
                    testdata[args.ProgressPercentage].nc1 = ""; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NC1 : Don't Care" );
                }
                if (no2 != "")
                {

                    Delay(dtime);
                    no2data = measvol(NO2add);
                    //listBox1.Items.Add("NO2 :" + no2data.ToString("#0.00"));
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NO2 :" + no2data.ToString("#0.00"));
                    if (no2data > 4.5) { no2act = "OFF"; }
                    if ((no2data < 2.4) && (no2data >= 0)) { no2act = "ON"; }
                    if ((no2data >= 2.4) && (no2data <= 4.5)) { no2act = "ERR"; }
                    if (no2data < 0) { no2act = "ERR"; }
                    testdata[args.ProgressPercentage].no2 = no2act;
                    dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                    if (no2 != no2act) { pass = 0; passitem = 0; passitem_no2=0; }
                    if (passitem_no2 != 0)
                    { testdata[args.ProgressPercentage].no2FColor = Brushes.Green; }
                    if (passitem_no2 == 0)
                    { testdata[args.ProgressPercentage].no2FColor = Brushes.Red; }


                }
                if (no2 == "")
                {
                    testdata[args.ProgressPercentage].no2 = ""; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NO2 : Don't Care");
                }

                if (nc3 != "")
                {
                    Delay(dtime);
                    nc3data = measvol(NC3add);
                   // listBox1.Items.Add("NC3 :" + nc3data.ToString("#0.00"));
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NC3 :" + nc3data.ToString("#0.00"));
                    if (nc3data > 4.5) { nc3act = "OFF"; }
                    if ((nc3data < 1) && (nc3data >= 0)) { nc3act = "ON"; }
                    if ((nc3data >= 1) && (nc3data <= 4.5)) { nc3act = "ERR"; }
                    if (nc3data < 0) { nc3act = "ERR"; }
                    testdata[args.ProgressPercentage].nc3 = nc3act;
                    dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                    if (nc3 != nc3act) { pass = 0; passitem = 0; passitem_nc3=0; }
                    if (passitem_nc3 != 0)
                    { testdata[args.ProgressPercentage].nc3FColor = Brushes.Green; }
                    if (passitem_nc3 == 0)
                    { testdata[args.ProgressPercentage].nc3FColor = Brushes.Red; }
                }
                if (nc3 == "") {
                    testdata[args.ProgressPercentage].nc3 = ""; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NC3 : Don't Care");
                }

                if (no4 != "")
                {
                    Delay(dtime);
                    //MessageBox.Show("33333");
                    no4data = measvol(NO4add);
                    //listBox1.Items.Add("NO4 :" + no4data.ToString("#0.00"));
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NO4 :" + no4data.ToString("#0.00"));
                    if (no4data > 4.5) { no4act = "OFF"; }
                    if ((no4data < 1) && (no4data >= 0)) { no4act = "ON"; }
                    if ((no4data >= 1) && (no4data <= 4.5)) { no4act = "ERR"; }
                    if (no4data < 0) { no4act = "ERR"; }
                    testdata[args.ProgressPercentage].no4 = no4act;
                    dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata;
                    if (no4 != no4act) { pass = 0; passitem = 0; passitem_no4=0; }
                    if (passitem_no4 != 0)
                    { testdata[args.ProgressPercentage].no4FColor = Brushes.Green; }
                    if (passitem_no4 == 0)
                    { testdata[args.ProgressPercentage].no4FColor = Brushes.Red; }
                }
                if (no4 == "") {
                    testdata[args.ProgressPercentage].no4 = ""; dataGridView1.ItemsSource = null;//这里是关键
                    dataGridView1.ItemsSource = testdata; Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "NO4 : Don't Care");
                }
            }//CoilNumber=="4"Zx
            if (passitem == 0) { testdata[args.ProgressPercentage].result = "FAIL"; testdata[args.ProgressPercentage].resultFColor = Brushes.Red; }
            if (passitem == 1) { testdata[args.ProgressPercentage].result = "PASS"; testdata[args.ProgressPercentage].resultFColor = Brushes.Green; }
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "");
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "");
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber2), "");
            flaggo = 0;
          //  MessageBox.Show(flaggo.ToString());














        }

       


        DataTable dt;
        private void RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs args)
        {
            // if (args.Cancelled) { MessageBox.Show("后台任务已经被取消。", "消息"); }
            //else { MessageBox.Show("后台任务正常结束。", "消息"); }
            if ((single == 0)&&(stopprocess==0))
            {
                test.IsEnabled = true;
                byte[] sendData;
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("VSET1:0");
                power1sp.Write(sendData, 0, sendData.Length);
                //    Delay(10); outport("Dev2/port0", 0x1F);
                sendcomm(sp, "ffaa030600000000b2");
                Delay(500);
                sendcomm(sp, "ffaa030800000000b4");
                if (pass == 0)
                {
                    this.Statue.Content = "FAIL";
                    this.Statue.Foreground = new SolidColorBrush(Colors.Red);
                   Lock pwprd = new Lock();  pwprd.ShowDialog();
                }
                if (pass == 1)
                {
                   /* outport("Dev2/port0", 0x13); Delay(300); outport("Dev2/port0", 0x1A); Delay(2500);*/ outport("Dev2/port0", 0x16); Delay(500); outport("Dev2/port0", 0x1E); Delay(50); outport("Dev2/port0", 0x1F); 
                    this.Statue.Content = "PASS";
                    this.Statue.Foreground = new SolidColorBrush(Colors.Green);
                }
                //savedata=savedata+
                for (int i = 0; i < Item; i++)
                {
                    if ((testdata[i].nc1 != "") && (testdata[i].nc1 != "None"))
                    { savedata = savedata + testdata[i].nc1 + ","; }
                    if ((testdata[i].no2 != "") && (testdata[i].no2 != "None"))
                    { savedata = savedata + testdata[i].no2 + ","; }
                    if ((testdata[i].nc3 != "") && (testdata[i].nc3 != "None"))
                    { savedata = savedata + testdata[i].nc3 + ","; }
                    if ((testdata[i].no4 != "") && (testdata[i].no4 != "None"))
                    { savedata = savedata + testdata[i].no4 + ","; }
                    if ((testdata[i].led != "") && (testdata[i].led != "None"))
                    { savedata = savedata + testdata[i].led + ","; }
                    if ((testdata[i].dcr1 != "") && (testdata[i].dcr1 != "None"))
                    { savedata = savedata + testdata[i].dcr1 + ","; }
                    if ((testdata[i].dcr2 != "") && (testdata[i].dcr2 != "None"))
                    { savedata = savedata + testdata[i].dcr2 + ","; }
                    if ((testdata[i].candata != "") && (testdata[i].candata != "None"))
                    { savedata = savedata + testdata[i].candata + ","; }


                }
                savefile = new FileStream(savepath, FileMode.Append);
                sw3 = new StreamWriter(savefile);
                if (pass == 1) { savedata = savedata + "PASS" + "\r\n"; }
                if (pass == 0)
                {
                    savedata = savedata + "FAIL" + "\r\n";
                }
                savedata = snno + "," + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "," + savedata;
                sw3.Write(savedata); sw3.Flush(); sw3.Close(); savefile.Close();
                con.Open();
                string sqlcomm = "use Danfossdata select Name from sysobjects where xtype='u'";
                SqlDataAdapter SqlDap = new SqlDataAdapter(sqlcomm, con);
                DataSet thisDataset = new DataSet();
                try
                {
                    SqlDap.Fill(thisDataset);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.ToString());

                }
             int xpan = 0,xpan2=0;
                string ss,tablename="D"+ MODEL.Content.ToString().Replace("-", "");
                do
                {
                    ss = thisDataset.Tables[0].Rows[xpan][0].ToString();
                    if (tablename == ss) { xpan2 = 1; }
                    xpan++;
                } while ((xpan< thisDataset.Tables[0].Rows.Count)&&(xpan2==0));
                string[] nc1item = new string[Item], no2item = new string[Item], nc3item = new string[Item], no4item = new string[Item];
                string[] leditem = new string[Item],dcr1item=new string[Item],dcr2item=new string[Item],canitem=new string[100];
                string[] cantotal;
                int cancount = 0;
                if (xpan2==0)
                {
                    dt = new DataTable("Table_Testdata");
                    dt.Columns.Add("model", System.Type.GetType("System.String"));
                    dt.Columns.Add("sn", System.Type.GetType("System.String"));
                    dt.Columns.Add("testtime", System.Type.GetType("System.String"));
                    DataRow dr = dt.NewRow();
                    dr[0] = MODEL.Content;
                    dr[1] = snno;
                    dr[2] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    int drcount = 3;
                    string sqlcomm2 = "USE Danfossdata" + "\r\n";
                    sqlcomm2 = sqlcomm2 + "CREATE TABLE D" + MODEL.Content.ToString().Replace("-","")+ "(Model varchar(50) NOT NULL, SN varchar(50) NOT NULL, Testtime varchar(50) NOT NULL,";
                    for (int i = 0; i < Item; i++)
                    {
                        
                        if ((testdata[i].nc1 != "") && (testdata[i].nc1 != "None"))
                        { nc1item[i] = "NC1" + i.ToString(); sqlcomm2 = sqlcomm2+ nc1item[i]+ " varchar(50),"; dt.Columns.Add(nc1item[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].nc1.ToString(); drcount++; }
                        if ((testdata[i].no2 != "") && (testdata[i].no2 != "None"))
                        { no2item[i] = "NO2" + i.ToString(); sqlcomm2 = sqlcomm2 + no2item[i] + " varchar(50),"; dt.Columns.Add(no2item[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].no2.ToString(); drcount++; }
                        if ((testdata[i].nc3 != "") && (testdata[i].nc3 != "None"))
                        { nc3item[i] = "NC3" + i.ToString(); sqlcomm2 = sqlcomm2 + nc3item[i] + " varchar(50),"; dt.Columns.Add(nc3item[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].nc3.ToString(); drcount++; }
                        if ((testdata[i].no4 != "") && (testdata[i].no4 != "None"))
                        { no4item[i] = "NO4" + i.ToString(); sqlcomm2 = sqlcomm2 + no4item[i] + " varchar(50),"; dt.Columns.Add(no4item[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].no4.ToString(); drcount++; }
                        if ((testdata[i].led != "") && (testdata[i].led != "None"))
                        { leditem[i] = "LED"+i.ToString(); sqlcomm2 = sqlcomm2 + leditem[i] + " varchar(50),"; dt.Columns.Add(leditem[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].led.ToString(); drcount++; }
                        if ((testdata[i].dcr1 != "") && (testdata[i].dcr1 != "None"))
                        { dcr1item[i] = "DCR1" + i.ToString(); sqlcomm2 = sqlcomm2 + dcr1item[i] + " varchar(50),"; dt.Columns.Add(dcr1item[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].dcr1.ToString(); drcount++; }
                        if ((testdata[i].dcr2 != "") && (testdata[i].dcr2 != "None"))
                        { dcr2item[i] = "DCR2" + i.ToString(); sqlcomm2 = sqlcomm2 + dcr2item[i] + " varchar(50),"; dt.Columns.Add(dcr2item[i], System.Type.GetType("System.String"));dr[drcount] = testdata[i].dcr2.ToString(); drcount++; }
                       
                        if ((testdata[i].candata != "") && (testdata[i].candata != "None"))
                        {
                            cantotal = testdata[i].candata.ToString().Split(',');
                            foreach (string azu in cantotal) { canitem[cancount] = "CanData" + cancount.ToString(); sqlcomm2 = sqlcomm2 + canitem[cancount] + " varchar(50),"; dt.Columns.Add(canitem[cancount], System.Type.GetType("System.String")); dr[drcount] = azu; drcount++; cancount++; }
                                
                        }


                    }
                  
                      // MessageBox.Show(dr[19].ToString());
                    dt.Columns.Add("Result", System.Type.GetType("System.String"));
                    dr[drcount] = pass.ToString();
                  //  MessageBox.Show("zuodeng");
                    dt.Rows.Add(dr);
                   // MessageBox.Show("zuodeng2");
                    sqlcomm2 = sqlcomm2 + "Result varchar(50) NOT NULL)";
                     //  MessageBox.Show(sqlcomm2);
                    cmd = new SqlCommand(sqlcomm2, con);
                    //执行sql命令
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.ToString());

                    }
                   
                    SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null);
                    bulkCopy.DestinationTableName = tablename;
                    bulkCopy.BatchSize = 1;
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        try
                        {
                            bulkCopy.WriteToServer(dt);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.ToString());

                        }
                    }
                    








                }
                if (xpan2 == 1)
                {
                    dt = new DataTable("Table_Testdata");
                    dt.Columns.Add("model", System.Type.GetType("System.String"));
                    dt.Columns.Add("sn", System.Type.GetType("System.String"));
                    dt.Columns.Add("testtime", System.Type.GetType("System.String"));
                    DataRow dr = dt.NewRow();
                    dr[0] = MODEL.Content;
                    dr[1] = snno;
                    dr[2] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    int drcount = 3;
                  
                    for (int i = 0; i < Item; i++)
                    {

                        if ((testdata[i].nc1 != "") && (testdata[i].nc1 != "None"))
                        {  dt.Columns.Add(nc1item[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].nc1.ToString(); drcount++; }
                        if ((testdata[i].no2 != "") && (testdata[i].no2 != "None"))
                        {  dt.Columns.Add(no2item[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].no2.ToString(); drcount++; }
                        if ((testdata[i].nc3 != "") && (testdata[i].nc3 != "None"))
                        { dt.Columns.Add(nc3item[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].nc3.ToString(); drcount++; }
                        if ((testdata[i].no4 != "") && (testdata[i].no4 != "None"))
                        {  dt.Columns.Add(no4item[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].no4.ToString(); drcount++; }
                        if ((testdata[i].led != "") && (testdata[i].led != "None"))
                        {  dt.Columns.Add(leditem[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].led.ToString(); drcount++; }
                        if ((testdata[i].dcr1 != "") && (testdata[i].dcr1 != "None"))
                        {  dt.Columns.Add(dcr1item[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].dcr1.ToString(); drcount++; }
                        if ((testdata[i].dcr2 != "") && (testdata[i].dcr2 != "None"))
                        { dt.Columns.Add(dcr2item[i], System.Type.GetType("System.String")); dr[drcount] = testdata[i].dcr2.ToString(); drcount++; }

                        if ((testdata[i].candata != "") && (testdata[i].candata != "None"))
                        {
                            cantotal = testdata[i].candata.ToString().Split(',');
                            foreach (string azu in cantotal) {dt.Columns.Add(canitem[cancount], System.Type.GetType("System.String")); dr[drcount] = azu; drcount++; cancount++; }

                        }


                    }

                    // MessageBox.Show(dr[19].ToString());
                    dt.Columns.Add("Result", System.Type.GetType("System.String"));
                    dr[drcount] = pass.ToString();
                    //  MessageBox.Show("zuodeng");
                    dt.Rows.Add(dr);
                    // MessageBox.Show("zuodeng2");                   
                    SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null);
                    bulkCopy.DestinationTableName = tablename;
                    bulkCopy.BatchSize = 1;
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        try
                        {
                            bulkCopy.WriteToServer(dt);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.ToString());

                        }
                    }

                }

                    //   MessageBox.Show(thisDataset.Tables[0].Rows[0][0].ToString());

                    /*  dt = new DataTable("Table_Testdata");
                    dt.Columns.Add("model", System.Type.GetType("System.String"));
                    dt.Columns.Add("sn", System.Type.GetType("System.String"));
                    dt.Columns.Add("testtime", System.Type.GetType("System.String"));
                    dt.Columns.Add("seq", System.Type.GetType("System.String"));
                    dt.Columns.Add("nc1", System.Type.GetType("System.String"));
                    dt.Columns.Add("no2", System.Type.GetType("System.String"));
                    dt.Columns.Add("nc3", System.Type.GetType("System.String"));
                    dt.Columns.Add("no4", System.Type.GetType("System.String"));
                    dt.Columns.Add("led", System.Type.GetType("System.String"));
                    dt.Columns.Add("dcr1", System.Type.GetType("System.String"));
                    dt.Columns.Add("dcr2", System.Type.GetType("System.String"));
                    dt.Columns.Add("candata", System.Type.GetType("System.String"));
                    dt.Columns.Add("result", System.Type.GetType("System.String"));
                    int d = Item,h;
                    SqlBulkCopy bulkCopy = new SqlBulkCopy(con, SqlBulkCopyOptions.TableLock, null);
                    bulkCopy.DestinationTableName = "Testdata";
                    bulkCopy.BatchSize = d;

                    for (int i = 0; i < Item; i++)
                    {
                        DataRow dr = dt.NewRow();
                        dr[0] = MODEL.Content;
                        dr[1] = snno;
                        dr[2] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        dr[3] = testdata[i].item.ToString();
                        dr[4] = testdata[i].nc1.ToString();
                        dr[5] = testdata[i].no2.ToString();
                        dr[6] = testdata[i].nc3.ToString();
                        dr[7] = testdata[i].no4.ToString();
                        dr[8] = testdata[i].led.ToString();
                        dr[9] = testdata[i].dcr1.ToString();
                        dr[10] = testdata[i].dcr2.ToString();
                        dr[11] = testdata[i].candata.ToString();
                        dr[12] = testdata[i].result.ToString();
                        dt.Rows.Add(dr);
                    }
                    if (dt != null && dt.Rows.Count != 0)
                    {
                        try
                        {
                            bulkCopy.WriteToServer(dt);
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.ToString());

                        }
                    }*/

                    con.Close();
                FileStream Debuginfo;
                StreamWriter sw4;
                //MessageBox.Show("d:\\" + snno + "-debuginfo.txt");
                string listbuf = "", listbuf2 = "";
                string snno1;
                snno1 = snno.Replace(":", "_");
                Debuginfo = new FileStream("d:\\" + snno1 + "-debuginfo.txt", FileMode.Create);
                sw4 = new StreamWriter(Debuginfo);
                for (int f = 0; f < list1.Items.Count; f++)
                {
                    listbuf2 = list1.Items[f].ToString();
                    listbuf = listbuf + listbuf2 + "\r\n";
                }
                sw4.Write(listbuf); sw4.Flush(); sw4.Close(); Debuginfo.Close();
                sw.Stop(); Time.Content = "Test time: " + sw.ElapsedMilliseconds.ToString() + "ms";
                if (CheckBox1.IsChecked == true)
                {
                    MouseButtonEventArgs keyargs = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left);
                    keyargs.RoutedEvent = Button.ClickEvent;
                    test.RaiseEvent(keyargs);
                }

            }

            if ((single == 0) && (stopprocess == 1))
            {
                test.IsEnabled = true;
                byte[] sendData;
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("VSET1:0");
                sendcomm(sp, "ffaa030600000000b2");
                Delay(500);
                power1sp.Write(sendData, 0, sendData.Length);
                    Delay(10); outport("Dev2/port0", 0x1F);
                sendcomm(sp, "ffaa030800000000b4");
            }


        }

        void Sn_TransfEvent(string value)
        {
            snno = value;
        }
        void Sn_TransfintEvent(int value)
        {
            snlength = value;
        }

        System.Diagnostics.Stopwatch sw ;

        private void ButtonEx_Click_1(object sender, RoutedEventArgs e)
        {
            // test.IsEnabled = false;
            single = 0;
            stopprocess = 0;
            teston = 1;
            if (pause == false)
            {
                test.IsEnabled = false;
                sw = System.Diagnostics.Stopwatch.StartNew();
              
                pass = 1; pass_item = 1;
                //  MessageBox.Show(ErrBox.Items.Count.ToString());
                //ErrBox.Items.Clear();           
                for (int f = ErrBox.Items.Count - 1; f > -1; f--)
                { _contentDescriptions.RemoveAt(f); }

                // MessageBox.Show("22");
                if (list1.Items.Count > 0)
                {
                    for (int f = list1.Items.Count - 1; f > -1; f--)
                    { _contentDescriptions2.RemoveAt(f); }
                }

                int i, x;
                i = dataGridView1.Items.Count;
                for (x = 0; x < i; x++)
                { testdata[x].candata = ""; testdata[x].dcr1 = ""; testdata[x].dcr2 = ""; testdata[x].led = ""; testdata[x].nc1 = ""; testdata[x].nc3 = ""; testdata[x].no2 = ""; testdata[x].no4 = ""; testdata[x].result = ""; }
                dataGridView1.ItemsSource = null;//这里是关键
                dataGridView1.ItemsSource = testdata;

                savedata = "";
                SN snbox = new SN();
                Regex regnum = new Regex("^[0~9]");
                int snpass = 1, index;
                string snmodel2, revx, rev2, days2, revr = "", pcbrev2;
                if (CheckBox1.IsChecked == true)
                {
                    snbox.TransfEvent += Sn_TransfEvent;
                    snbox.TransintfEvent += Sn_TransfintEvent;
                    snbox.ShowDialog();
                    snbox2.Text = snno;
                    if (snlength == 0) {
                        teston = 0; this.Statue.Content = "SN LENGTH 0";
                        this.Statue.Foreground = new SolidColorBrush(Colors.Red);
                        test.IsEnabled = true;
                    }
                    if (snlength != 0)
                    {
                        if (snlength != samplelength) { this.Statue.Content = "Length Err"; this.Statue.Foreground = new SolidColorBrush(Colors.Red); teston = 0; snpass = 0; test.IsEnabled = true; }
                        if (snpass == 1)
                        {
                            index = snno.IndexOf(":");
                            snmodel2 = snno.Substring(0, index);
                            if (snmodel2 != smodel)
                            {
                                this.Statue.Content = "Model Err";
                                this.Statue.Foreground = new SolidColorBrush(Colors.Red); // 颜色 
                                snpass = 0; teston = 0; test.IsEnabled = true;
                            }
                        }
                        if (snpass == 1)
                        {
                            revx = getrev(snno, revr);
                            //  MessageBox.Show(rev);  
                            if (regnum.IsMatch(revx.Substring(0, 1)))
                            {
                                // MessageBox.Show("Num");
                                rev2 = revx.Substring(0, 2);
                                days2 = revx.Substring(2, 4);
                                // MessageBox.Show(rev);
                            }
                            else
                            {
                                rev2 = revx.Substring(0, 1);
                                days2 = revx.Substring(1, 4);
                                // MessageBox.Show(rev);
                            }
                            if (rev2 != rev)
                            {
                                this.Statue.Content = "rev Err";
                                this.Statue.Foreground = new SolidColorBrush(Colors.Red);// 颜色 
                                snpass = 0; teston = 0; test.IsEnabled = true;
                            }
                            if (days2 != pdays)
                            {
                                this.Statue.Content = "day Err";
                                this.Statue.Foreground = new SolidColorBrush(Colors.Red);// 颜色 
                                snpass = 0; teston = 0; test.IsEnabled = true;
                            }
                        }
                        if (snpass == 1)
                        {
                            revr = "";
                            pcbrev2 = getpcbrev(snno, revr);
                            if (pcbrev2 != pcbrev)
                            {
                                this.Statue.Content = "PCBREV Err";
                                this.Statue.Foreground = new SolidColorBrush(Colors.Red);// 颜色 
                                snpass = 0; teston = 0; test.IsEnabled = true;
                            }
                        }


                    }





                }

                if (CheckBox1.IsChecked == false) { teston = 1; }

                if (teston == 1)
                {
                    if ((MODEL.Content.ToString().IndexOf("-90") == -1)&& (MODEL.Content.ToString().IndexOf("0123") == -1)) { outport("Dev2/port0", 0x1B);Delay(500); outport("Dev2/port0", 0x1A); }
                    if ((MODEL.Content.ToString().IndexOf("-90") == -1) && (MODEL.Content.ToString().IndexOf("0123") != -1)) { outport("Dev3/port0", 0x1A); Delay(1000); outport("Dev2/port0", 0x17); Delay(500); outport("Dev2/port0", 0x1A); }
                        Delay(300);
                    this.Statue.Content = "TESTING";
                    this.Statue.Foreground = new SolidColorBrush(Colors.Blue);
                    if (!bgWorker.IsBusy)

                    {

                        bgWorker.RunWorkerAsync();

                    }
                }
            }
            if(pause==true)
            {
                test.IsEnabled = false;
                pause = false;
                manualReset.Set();//
            }

        }

     

        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "CreatTask", CallingConvention = CallingConvention.Cdecl)]
        public static extern UInt32 CreatTask();
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "GetErrInfo", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern IntPtr GetErr(Int32 code);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "ReadDCVol", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern Int32 GetDCVol(IntPtr taskhandel, Int32 numSampsPerChan, double timeout, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] double[] readArray, out Int32 sampsPerChanRead);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "Writeport", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern Int32 Writeport(IntPtr taskhandel, Int32 numSampsPerChan, UInt32 autostart, double timeout, UInt32 data, out Int32 sampsPerChanWrite);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "ConfigChann", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern Int32 ConfigChann(IntPtr taskhandel, [MarshalAs(UnmanagedType.LPStr)]string param1, [MarshalAs(UnmanagedType.LPStr)]string param2, Int32 terminalConfig, double minVal, double maxVal, Int32 units);
        [DllImport(@"D:\nidaqdll.dll", EntryPoint = "ConfigDOChann", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Winapi)]
        public static extern Int32 ConfigDOChann(IntPtr taskhandel, [MarshalAs(UnmanagedType.LPStr)]string param1, [MarshalAs(UnmanagedType.LPStr)]string param2);
        [DllImport("kernel32")] //返回0表示失败，非0为成功
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")] //返回取得字符串缓冲区的长度
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        private void Equset_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show("Equ");
            Equipment equ = new Equipment();
            equ.ShowDialog();
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendData;
            string errinfo;
            sendData = null;
            sendData = Encoding.UTF8.GetBytes("VSET1:0");
            power1sp.Write(sendData, 0, sendData.Length); ; Delay(10); outport("Dev2/port0", 0x1F);
            errinfo = sendcomm(sp, "ffaa030800000000b4");
            if (errinfo != "")
            {
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errinfo);
            }
        }
        int stopprocess=0;
        private void STOP_Click(object sender, RoutedEventArgs e)
        {
            stopprocess = 1;
            this.Statue.Content = "STOP";
            this.Statue.Foreground = new SolidColorBrush(Colors.Red);
            test.IsEnabled = true;
            if (single == 0) { bgWorker.CancelAsync(); }
            if (single != 0)
            {
                byte[] sendData;
                sendData = null;
                sendData = Encoding.UTF8.GetBytes("VSET1:0");
                power1sp.Write(sendData, 0, sendData.Length) ; Delay(10); outport("Dev2/port0", 0x1F);
                sendcomm(sp, "ffaa030600000000b2");
                Delay(500);
                sendcomm(sp, "ffaa030800000000b4");

            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            /* if ((e.KeyboardDevice.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
             {
                 MessageBox.Show("您按下了Control键");
             }
             if (e.KeyboardDevice.IsKeyDown(Key.N))
             { MessageBox.Show("您按下了N键"); }*/
            if (e.KeyboardDevice.IsKeyDown(Key.Enter))
            {
                if (test.IsEnabled == true)
                {
                    MouseButtonEventArgs args = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left);
                    args.RoutedEvent = Button.ClickEvent;
                    test.RaiseEvent(args);
                }
                }
            if (e.KeyboardDevice.IsKeyDown(Key.Escape))
            {
               // if (test.IsEnabled == true)
               // {
                    MouseButtonEventArgs args = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left);
                    args.RoutedEvent = Button.ClickEvent;
                    STOP.RaiseEvent(args);
               // }
            }


        }

        private void Window_Closed(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }
        private ManualResetEvent manualReset = new ManualResetEvent(true);
        
        private void Pause_Click(object sender, RoutedEventArgs e)
        {
            test.IsEnabled = true;
            pause = true;
            manualReset.Reset();//
        }
        private int singleitem;
        private void Si_Click(object sender, RoutedEventArgs e)
        {
            single = 1;
            int i, x;
            i = dataGridView1.SelectedIndex;
            if (list1.Items.Count > 0)
            {
                for (int f = list1.Items.Count - 1; f > -1; f--)
                { _contentDescriptions2.RemoveAt(f); }
            }
            
          //  for (x = 0; x < i; x++)
             testdata[i].candata = ""; testdata[i].dcr1 = ""; testdata[i].dcr2 = ""; testdata[i].led = ""; testdata[i].nc1 = ""; testdata[i].nc3 = ""; testdata[i].no2 = ""; testdata[i].no4 = ""; testdata[i].result = ""; 
            dataGridView1.ItemsSource = null;//这里是关键
            dataGridView1.ItemsSource = testdata;

            if ((MODEL.Content.ToString().IndexOf("-90") == -1)&& (MODEL.Content.ToString().IndexOf("-0123") == -1)) { outport("Dev2/port0", 0x1A); }
            if ((MODEL.Content.ToString().IndexOf("-90") == -1) && (MODEL.Content.ToString().IndexOf("-0123") != -1)) { outport("Dev3/port0", 0x1A); Delay(1000); outport("Dev2/port0", 0x1A); }
            singleitem = i;
            //MessageBox.Show(singleitem.ToString());
            if (!bgWorker.IsBusy)
            {
                bgWorker.RunWorkerAsync();
            }
        }

        private void StopMotor_Click(object sender, RoutedEventArgs e)
        {
            sendcomm(sp, "ffaa030600000000b2");
            Delay(500);
        }

        private void DataExpert_Click(object sender, RoutedEventArgs e)
        {
            Expert exp = new Expert();
            exp.modelset = MODEL.Content.ToString();
            exp.itemSet = Item;
            exp.ShowDialog();
        }

        public class Testdata
        {
            public string seq { get; set; }
            public string item { get; set; }
            //  public string nc1 { get; set; }
            private string _nc1;
            public string nc1
            {
                get { return _nc1; }
                set { _nc1 = value; }
            }
            private SolidColorBrush _nc1FColor;
            public SolidColorBrush nc1FColor
            {
                get { return _nc1FColor; }
                set { _nc1FColor = value; }
            }
            //    public string no2 { get; set; }
            private string _no2;
            public string no2
            {
                get { return _no2; }
                set { _no2 = value; }
            }
            private SolidColorBrush _no2FColor;
            public SolidColorBrush no2FColor
            {
                get { return _no2FColor; }
                set { _no2FColor = value; }
            }
            //    public string nc3 { get; set; }
            private string _nc3;
            public string nc3
            {
                get { return _nc3; }
                set { _nc3 = value; }
            }
            private SolidColorBrush _nc3FColor;
            public SolidColorBrush nc3FColor
            {
                get { return _nc3FColor; }
                set { _nc3FColor = value; }
            }


            // public string no4 { get; set; }
            private string _no4;
            public string no4
            {
                get { return _no4; }
                set { _no4 = value; }
            }
            private SolidColorBrush _no4FColor;
            public SolidColorBrush no4FColor
            {
                get { return _no4FColor; }
                set { _no4FColor = value; }
            }




            //public string led { get; set; }
            private string _led;
            public string led
            {
                get { return _led; }
                set { _led = value; }
            }
            private SolidColorBrush _ledFColor;
            public SolidColorBrush ledFColor
            {
                get { return _ledFColor; }
                set { _ledFColor = value; }
            }


            //  public string dcr1 { get; set; }
            private string _dcr1;
            public string dcr1
            {
                get { return _dcr1; }
                set { _dcr1 = value; }
            }
            private SolidColorBrush _dcr1FColor;
            public SolidColorBrush dcr1FColor
            {
                get { return _dcr1FColor; }
                set { _dcr1FColor = value; }
            }

            // public string dcr2 { get; set; }
            private string _dcr2;
            public string dcr2
            {
                get { return _dcr2; }
                set { _dcr2 = value; }
            }
            private SolidColorBrush _dcr2FColor;
            public SolidColorBrush dcr2FColor
            {
                get { return _dcr2FColor; }
                set { _dcr2FColor = value; }
            }
            // public string candata { get; set; }
            private string _candata;
            public string candata
            {
                get { return _candata; }
                set { _candata = value; }
            }
            private SolidColorBrush _candataFColor;
            public SolidColorBrush candataFColor
            {
                get { return _candataFColor; }
                set { _candataFColor = value; }
            }

            //   public string result { get; set; }
            private string _result;
            public string result
            {
                get { return _result; }
                set { _result = value; }
            }
            private SolidColorBrush _resultFColor;
            public SolidColorBrush resultFColor
            {
                get { return _resultFColor; }
                set { _resultFColor = value; }
            }

        }

        string content = "Start";
        private void Open1_Click(object sender, RoutedEventArgs e)
        {
            int rw;
            rw = testdata.Count();
            if (rw > 0)
            {
                for(int fc = rw; fc >0; fc--) { testdata.RemoveAt(fc-1); }
            }
            /* dataGridView1.ItemsSource = null;
             dataGridView1.ItemsSource = testdata;
           */
            //open1.IsChecked = true;
            
            Item = 0;
            open = 1;
            ofile = ofile + 1;
            StringBuilder temp = new StringBuilder(500);
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            string motorstr, errinfo, model;
            string[] seq;
            ofd.Filter = "ini文件(*.ini;*.txt)|*.ini;*.txt|所有文件|*.*";
            ofd.ValidateNames = true;
            ofd.CheckPathExists = true;
            ofd.CheckFileExists = true;
            if (ofd.ShowDialog() == true)
            {
              //  testdata = null;
                ff = ofd.FileName;
                GetPrivateProfileString("Info", "TestItem", "3", temp, 500, ff);
                seq = temp.ToString().Split(',');
                int x2 = 0, x = 0;
                string item = "Item";
                foreach (string azu in seq)
                { x++; }
               Seq = new int[x];
                foreach (string azu in seq)
                {
                    Seq[x2] = int.Parse(azu);
                    item = item + Seq[x2].ToString();
                    GetPrivateProfileString(item, "Item", "N/A", temp, 500, ff);
                    testdata.Add(new Testdata
                    { seq = (x2 + 1).ToString(),item=temp.ToString()  }
                        );
                    dataGridView1.DataContext = testdata;
                    item = "Item";
                    x2++;
                    Item++;
                }
               
                //  testdata.RemoveAt(0);
                //   errcontent2 = "test2";
                //  Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber),"112");
                
                GetPrivateProfileString("HardInfo", "StepMotor", "0", temp, 500, "D:\\HardInfo.ini");
                motorstr = temp.ToString();
                motorstr = motorstr.Replace("ASRL", "com");
                StopMotor.IsEnabled = true;
                if (ofile == 1)
                {
                  //  MessageBox.Show(motorstr.Substring(0, 4));
                    sp = new SerialPort(motorstr.Substring(0, 4));
                    sp.BaudRate = 9600;
                    sp.DataBits = 8;
                    sp.Parity = Parity.None;
                    sp.StopBits = StopBits.Two;
                    sp.Handshake = Handshake.None;
                    try
                    {

                        sp.Open();

                    }

                    catch (Exception ee)
                    {
                        errcontent2 = "StepMotor Com Err:" + ee.Message;
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                        //ErrBox.Items.Add("StepMotor Com Err:" + ee);
                        //sp.Close();
                        //sp.Open();
                    }
                    sp.ReceivedBytesThreshold = 1;
                    sp.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(sp_DataReceived);
                }
              //  MessageBox.Show("222");
                errinfo = sendcomm(sp, "FFAA03019001B400F2");
                if (errinfo != "")
               // { ErrBox.Items.Add("StepMotor Com Err:" + errinfo); }
                {
                    errcontent2 = "StepMotor Com Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                }
                errinfo = sendcomm(sp, "FFAA03059600010048");
                if (errinfo != "")
                {
                    errcontent2 = "StepMotor Com Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                }
                errinfo = sendcomm(sp, "FFAA030C01000000B9");
                if (errinfo != "")
                //{ ErrBox.Items.Add("StepMotor Com Err:" + errinfo); }
                {
                    errcontent2 = "StepMotor Com Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                }
                errinfo = sendcomm(sp, "FFAA030A00000000B6");
                if (errinfo != "")
                // { ErrBox.Items.Add("StepMotor Com Err:" + errinfo); }
                {
                    errcontent2 = "StepMotor Com Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                }
                errinfo = sendcomm(sp, "FFAA030E00000000BA");
                if (errinfo != "")
               // { ErrBox.Items.Add("StepMotor Com Err:" + errinfo); }
                {
                    errcontent2 = "StepMotor Com Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                }
                errinfo = sendcomm(sp, "FFAA030800000000B4");
                if (errinfo != "")
                //  { ErrBox.Items.Add("StepMotor Com Err:" + errinfo); }
                {
                    errcontent2 = "StepMotor Com Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                }
                sendcomm(sp, "ffaa030600000000b2");
                
                GetPrivateProfileString("HardInfo", "Keithley", "0", temp, 500, "D:\\HardInfo.ini");
                keithleyadd = temp.ToString();
                GetPrivateProfileString("HardInfo", "Pow1", "0", temp, 500, "D:\\HardInfo.ini");
                pow1add = temp.ToString();
                if (pow1add != "0")
                {
                    pow1add = pow1add.Replace("ASRL", "com");
                    if (ofile == 1)
                    {
                        power1sp = new SerialPort(pow1add.Substring(0, 4));
                        power1sp.BaudRate = 9600;
                        power1sp.DataBits = 8;
                        power1sp.Parity = Parity.None;
                        power1sp.StopBits = StopBits.Two;
                        power1sp.Handshake = Handshake.None;
                        try
                        {

                            power1sp.Open();

                        }

                        catch (Exception ee)
                        {
                            // ErrBox.Items.Add("Power1 Com Err:" + ee);
                            errcontent2 = "Power1 Com Err:" + ee.Message;
                            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                            //       power1sp.Close();
                            //     power1sp.Open();

                        }
                    }

                    power1sp.ReceivedBytesThreshold = 1;
                    power1sp.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(sp2_DataReceived);
                }
                GetPrivateProfileString("HardInfo", "Pow2", "0", temp, 500, "D:\\HardInfo.ini");
                pow2add = temp.ToString();
                if (pow2add != "0")
                {
                    pow2add = pow2add.Replace("ASRL", "com");
                    if (ofile == 1)
                    {
                        power2sp = new SerialPort(pow2add.Substring(0, 4));
                        power2sp.BaudRate = 9600;
                        power2sp.DataBits = 8;
                        power2sp.Parity = Parity.None;
                        power2sp.StopBits = StopBits.One;
                        power2sp.Handshake = Handshake.None;
                        try
                        {

                            power2sp.Open();

                        }

                        catch (Exception ee)
                        {
                            //ErrBox.Items.Add("Power2 Com Err:" + ee);
                            errcontent2 = "Power2 Com Err:" + ee.Message;
                            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                            //           power2sp.Close();
                            //         power2sp.Open();
                        }
                    }
                    power2sp.ReceivedBytesThreshold = 1;
                    power2sp.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(sp3_DataReceived);
                }
                GetPrivateProfileString("HardInfo", "NC1", "0", temp, 500, "D:\\HardInfo.ini");
                NC1add = temp.ToString();
                GetPrivateProfileString("HardInfo", "NO2", "0", temp, 500, "D:\\HardInfo.ini");
                NO2add = temp.ToString();
                GetPrivateProfileString("HardInfo", "NC3", "0", temp, 500, "D:\\HardInfo.ini");
                NC3add = temp.ToString();
                GetPrivateProfileString("HardInfo", "NO4", "0", temp, 500, "D:\\HardInfo.ini");
                NO4add = temp.ToString();
                GetPrivateProfileString("HardInfo", "LED", "0", temp, 500, "D:\\HardInfo.ini");
                LEDadd = temp.ToString();
                GetPrivateProfileString("HardInfo", "Keithley", "0", temp, 500, "D:\\HardInfo.ini");
                DCRADD = temp.ToString();
                errinfo = outport("Dev2/port0", 0x1F);
                if (errinfo != "")
                //{ ErrBox.Items.Add("Daqoutport Err:" + errinfo); }
                {
                    errcontent2 = "Daqoutport Err:" + errinfo;
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                    }
                GetPrivateProfileString("Info", "CanEnable", "0", temp, 500, ff);
                CanEnable = temp.ToString();
                if (int.Parse(CanEnable) == 1)
                {

                    if (Ecan.OpenDevice(4, 0, 0) != Ecan.ECANStatus.STATUS_OK)
                  //  { ErrBox.Items.Add("Open CanDevice Err"); }
                    {
                        errcontent2 = "Open CanDevice Err";
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                         }
                    Ecan.INIT_CONFIG init_config = new Ecan.INIT_CONFIG();
                    init_config.AccCode = 0;
                    init_config.AccMask = 0xffffffff;
                    init_config.Filter = 0;
                    init_config.Timing0 = 0x01;
                    init_config.Timing1 = 0x1C;
                    init_config.Mode = 0;
                    init_config.Reserved = 0x00;
                    if (Ecan.InitCAN(4, 0, 0, ref init_config) != Ecan.ECANStatus.STATUS_OK)
                  //  { ErrBox.Items.Add("Init Can0 fault!"); }
                    {
                        errcontent2 = "Init Can0 Fail!";
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                         }
                    if (Ecan.InitCAN(4, 0, 1, ref init_config) != Ecan.ECANStatus.STATUS_OK)
                   // { ErrBox.Items.Add("Init Can1 fault!"); }
                    {
                        errcontent2 = "Init Can1 Fail!";
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                        }
                    if (Ecan.StartCAN(4, 0, 0) != Ecan.ECANStatus.STATUS_OK)
                    {
                        errcontent2 = "Start Can0 Fail!";
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);
                       }
                    if (Ecan.StartCAN(4, 0, 1) != Ecan.ECANStatus.STATUS_OK)
                    {
                        errcontent2 = "Start Can1 Fail!";
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);               
                    }
                      //  ErrBox.Items.Add("Start Can1 Fail!"); }
                }
                GetPrivateProfileString("Info", "CoilNumber", "0", temp, 500, ff);
                CoilNumber = temp.ToString();
                GetPrivateProfileString("Info", "Model", "Model", temp, 500, ff);
                model = temp.ToString();
                MODEL.Content = model;
                path2 = DateTime.Now.ToString("yyyy");
                savepath = "E:\\DanfossTEST";
                savepath = savepath + "\\" + model;
               // MessageBox.Show(savepath);
                if (!Directory.Exists(savepath))
                {
                    Directory.CreateDirectory(savepath);
                }
                savepath = savepath + "\\" + path2 + ".txt";
               // MessageBox.Show(savepath);
                if (!File.Exists(savepath))
                {
                    savefile = new FileStream(savepath, FileMode.Create);
                    sw3 = new StreamWriter(savefile);
                    filenew = true;
                    sw3.Flush();
                    //关闭流
                    sw3.Close();
                    savefile.Close();
                }
                reset.IsEnabled = true;
                test.IsEnabled = true;
                test.Focus();
                DataExpert.IsEnabled = true;











            }


        }

           private List<Model> Setcontent()
           {
               List<Model> listcontent = new List<Model>();

                   listcontent.Add(new Model() { errcontent = content });

               return listcontent;
           }

       public class Model
       {
           public string errcontent { get; set; }
       }
          // void LoadData()
          // { ErrBox.ItemsSource = Setcontent(); }
   
        public class ViewModel : INotifyPropertyChanged
        {

            public event PropertyChangedEventHandler PropertyChanged;
            public void OnPropertyChanged(PropertyChangedEventArgs e)
            {
                if (PropertyChanged != null)
                    PropertyChanged(this, e);
            }


            private ObservableCollection<string> listStatus = new ObservableCollection<string> { };
            public ObservableCollection<string> ListStatus
            {
                get { return listStatus; }
                set
                {
                    listStatus = value;
                    OnPropertyChanged(new PropertyChangedEventArgs("ListStatus"));
                }
            }

        }
        string errcontent2;
        private ObservableCollection<string> _contentDescriptions;
        private ObservableCollection<string> _contentDescriptions2;
        private void LoadNumber(string content)
        {
           // content = errcontent2;
            _contentDescriptions.Add(content);
           // MessageBox.Show("33");
          //  Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber));
        }
        private void LoadNumber2(string content)
        {
            // content = errcontent2;
            _contentDescriptions2.Add(content);
            // MessageBox.Show("33");
            //  Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber));
        }
        private delegate void LoadContentDelegate(string content);




        private void sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Byte[] result = new Byte[6];
            string resultbuff, result2 = "";
            int i;
            resultbuff = "";
            sp.Read(result, 0, 6);
            //  result2 = power1sp.ReadLine();
            //  MessageBox.Show(result.ToString());
            if (result != null)
            {
                for (i = 0; i < result.Length; i++)
                { resultbuff += result[i].ToString("X2"); }
            }
            //ViewModel errcontext = new ViewModel();
            // errcontext.ListStatus.Add(new ViewModel { listStatus = ""; });
            errcontent2 = "Step Common:" + resultbuff;
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber),errcontent2);
            //  ErrBox.Items.Add("Step Common:" + resultbuff);
            // ErrBox.ItemsSource = null;
            //  content = "Step Common:" + resultbuff;
            // ErrBox.ItemsSource = Setcontent();
            //  Setcontent();
            //   LoadData();
        }
        private void sp2_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //  Byte[] result = new Byte[6];
            string resultbuff;
            int i;
            Byte[] receivedData = new Byte[power1sp.BytesToRead];        //创建接收字节数组
            power1sp.Read(receivedData, 0, receivedData.Length);         //读取数据
            resultbuff = new UTF8Encoding().GetString(receivedData);//
                                                                    /* resultbuff = "";
                                                                     power1sp.Read(result, 0, 6);
                                                                     if (result != null)
                                                                     {
                                                                         for (i = 0; i < result.Length; i++)
                                                                         { resultbuff += result[i].ToString("X2"); }
                                                                     }*/
            errcontent2 = "Vbat:" + resultbuff;
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);                                                       // if (result2!="")
                                                             // {
           // ErrBox.Items.Add("Vbat:" + resultbuff);
            //  }


        }
        private void sp3_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //  Byte[] result = new Byte[6];
            string resultbuff;
            int i;
            Byte[] receivedData = new Byte[power2sp.BytesToRead];        //创建接收字节数组
            power2sp.Read(receivedData, 0, receivedData.Length);         //读取数据
            resultbuff = new UTF8Encoding().GetString(receivedData);//
                                                                    /* resultbuff = "";
                                                                     power1sp.Read(result, 0, 6);
                                                                     if (result != null)
                                                                     {
                                                                         for (i = 0; i < result.Length; i++)
                                                                         { resultbuff += result[i].ToString("X2"); }
                                                                     }*/
                                                                    // if (result2!="")
            errcontent2 = "Vbat2:" + resultbuff;
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadContentDelegate(LoadNumber), errcontent2);                                             // {
            //ErrBox.Items.Add("Vbat2:" + resultbuff);
            //  }


        }
        string motor_Dis(string disstr)
        {
            string dis = "";
            int pulse, i;
            string pbuff2, comm, checksum;
            string[] pbuff;
            Int32[] sum;
            Int32 sum2;
            sum2 = 0;
            pbuff = new string[3];
            sum = new Int32[8];
            pulse = Int32.Parse(disstr);
            pbuff2 = pulse.ToString("X");
            if (pulse.ToString("X").Length < 6)
            {
                for (i = 0; i < 6 - pulse.ToString("X").Length; i++)
                { pbuff2 = "0" + pbuff2; }
            }
            for (i = 0; i < 3; i++) { pbuff[i] = pbuff2.Substring(i * 2, 2); }
            comm = "ffaa0303" + pbuff[2] + pbuff[1] + pbuff[0] + "00";
            byte[] ee3 = new byte[comm.Length / 2];
            for (i = 0; i < ee3.Length; i++)
            {
                ee3[i] = Convert.ToByte(comm.Substring(i * 2, 2), 16);
            }
            for (i = 0; i < ee3.Length; i++) { sum[i] = Convert.ToInt32(ee3[i].ToString("X"), 16); }
            for (i = 0; i < ee3.Length; i++) { sum2 = sum2 + sum[i]; }
            if (sum2.ToString("X").Length >= 2) { checksum = sum2.ToString("X").Substring(sum2.ToString("X").Length - 2, 2); }
            else { checksum = "0" + sum2.ToString("X"); }
            comm = comm + checksum; dis = comm;
            return dis;
        }

        private int hextodec(string ss)//hex transform bin then bin transform dec can transform negative number
        {
            int i, shi, actdata = 0;
            long ww;
            string er1;
            string[] er;
            int leng, ix = 0;
            string erx = "";
            string erx2, erx4 = "";
            string[] erx3;
            erx3 = new string[16];
            er = new string[16];
            ww = Int32.Parse(ss, System.Globalization.NumberStyles.HexNumber);
            er1 = Convert.ToString(ww, 2);
            leng = er1.Length;
            if (leng < 16)
            {
                for (i = 0; i < 16 - leng; i++)
                {
                    er[i] = "0";
                }


                for (i = 16 - leng; i < 16; i++)
                {
                    er[i] = er1.Substring(ix, 1);
                    ix = ix + 1;
                    // MessageBox.Show(er[i]);
                }
                // MessageBox.Show("33");
            }
            else
            {

                for (i = 0; i < 16; i++)
                {
                    er[i] = er1.Substring(i, 1);
                }
            }

            if (er[0] == "1")
            {
                for (i = 0; i < 16; i++)
                {
                    erx = erx + er[i];
                }
                // MessageBox.Show(erx);
                shi = Convert.ToInt32(erx, 2);
                shi = shi - 1;
                erx2 = Convert.ToString(shi, 2);
                for (i = 0; i < 16; i++)
                {
                    if (erx2.Substring(i, 1) == "1")
                    {
                        erx3[i] = "0";
                    }
                    if (erx2.Substring(i, 1) == "0")
                    {
                        erx3[i] = "1";
                    }
                }
                for (i = 0; i < 16; i++)
                {
                    erx4 = erx4 + erx3[i];
                }
                actdata = Convert.ToInt32(erx4, 2);
                actdata = 0 - actdata;


            }
            if (er[0] == "0")
            {
                for (i = 0; i < 16; i++)
                {
                    erx = erx + er[i];
                }
                actdata = Convert.ToInt32(erx, 2);
            }
            return actdata;
        }



        SerialPort sp, power1sp, power2sp;
        FileStream savefile;
        StreamWriter sw3;
        int[] Seq;
        string savedata, canmodel = "";
      
        int samplelength, teston = 1, snlength, Item = 0, pass = 1, pass_item = 1, stop = 0, u = 0, open = 0, signal = 0, ofile = 0, actframe = 0;
        bool flag = true, filenew, pause = false;








    }
}
