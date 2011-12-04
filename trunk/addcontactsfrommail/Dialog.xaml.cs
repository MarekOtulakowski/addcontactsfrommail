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

namespace addcontactsfrommail
{
    /// <summary>
    /// Interaction logic for Dialog.xaml
    /// </summary>
    public partial class Dialog : Window
    {
        private string _windowName;
        private string _question;

        public Dialog()
        {
            InitializeComponent();
        }

        public Dialog(string windowName,
                      string question)
        {
            InitializeComponent();

            SolidColorBrush myBrush = new SolidColorBrush(Colors.Black);
            this.Background = myBrush;
            this.Title = windowName;
            this.L_question.Content = (object)question;

            //_windowName = windowName;
            //_question = windowName;
            //Initialize();
        }

        //private void Initialize()
        //{
        //    this.Title = _windowName;
        //    this.L_question.Content = _question;
        //}

        private void B_yes_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }

        private void B_no_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
    }
}
