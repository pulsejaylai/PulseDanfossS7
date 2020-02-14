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
    /// Interaction logic for SN.xaml
    /// </summary>
    public partial class SN : Window
    {
        public delegate void TransfDelegate(String value);
        public delegate void TransintfDelegate(int valuet);
        public event TransfDelegate TransfEvent;
        public event TransintfDelegate TransintfEvent;
        public SN()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            sn.Text = "";
            sn.Focus();
        }

        private void Sn_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyboardDevice.IsKeyDown(Key.Enter))
            {
                
                
                    MouseButtonEventArgs args = new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left);
                    args.RoutedEvent = Button.ClickEvent;
                    SNOK.RaiseEvent(args);
                
            }
        }

        private void SNOK_Click(object sender, RoutedEventArgs e)
        {
            int length;           
            length = sn.Text.Length;
            TransfEvent(sn.Text);
            TransintfEvent(length);
            this.Close();
        }


    }
}
