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

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for SampleSN.xaml
    /// </summary>
    public partial class SampleSN : Window
    {
        public SampleSN()
        {
            InitializeComponent();
        }
        public delegate void TransfDelegate(String value);
        public delegate void TransintfDelegate(int valuet);
        public event TransfDelegate TransfEvent;
        public event TransintfDelegate TransintfEvent;

        private void SMSN_Click(object sender, RoutedEventArgs e)
        {
            int length;
            length = SMSNINPUT.Text.Length;
            TransfEvent(SMSNINPUT.Text);
            TransintfEvent(length);
            this.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SMSNINPUT.Text = "";
            SMSNINPUT.Focus();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyboardDevice.IsKeyDown(Key.Enter))
            {


                MouseButtonEventArgs args = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left);
                args.RoutedEvent = Button.ClickEvent;
                SMSN.RaiseEvent(args);

            }
        }


    }
}
