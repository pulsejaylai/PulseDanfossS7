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
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Threading;
using System.Windows.Threading;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for Lock.xaml
    /// </summary>
    public partial class Lock : Window
    {
        public Lock()
        {
            InitializeComponent();
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

        int pas = 0;
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)

        {
            if (pas == 0)
            { e.Cancel = true; }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private void TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyboardDevice.IsKeyDown(Key.Enter))
            {


                MouseButtonEventArgs args = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left);
                args.RoutedEvent = Button.ClickEvent;
                button1.RaiseEvent(args);

            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            string checkword;
            FileStream wordtext, checktext;
            wordtext = new FileStream("C:\\Windows\\System\\password.txt", FileMode.Open);
            StreamReader sr3 = new StreamReader(wordtext); checkword = sr3.ReadLine(); wordtext.Close();
            if (checkword == textBox1.Password)
            {
                checktext = new FileStream("C:\\Windows\\System32\\dcheck.txt", FileMode.Create);
                StreamWriter checkok = new StreamWriter(checktext); checkok.Write("PASS"); checkok.Flush(); checkok.Close(); checktext.Close();
                outport("Dev2/port0", 0x1B); Delay(30); outport("Dev2/port0", 0x1F); 
                pas = 1;
                this.Close();


            }
            if (checkword != textBox1.Password)
            {
                MessageBox.Show("PassWord Err");
            }
        }

        public delegate string Outport1(string Chan, UInt32 Data);
        Outport1 outport = (Chan, Data) =>
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            textBox1.Focus();
               
        }





    }
}
