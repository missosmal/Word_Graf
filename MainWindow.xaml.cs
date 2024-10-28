using System.Windows;
using Word_Graf.Context;
using Microsoft.Win32;

namespace Word_Graf
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadRooms();
        }

        private void Report(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Files (*.docx)|*.docx";
            sfd.ShowDialog();
            if (sfd.FileName != "")
                OwnerContext.Report(sfd.FileName);
        }
        public void LoadRooms()
        {
            for (int i = 1; i < 20; i++)
                Parent.Children.Add(new Elements.Room(i));
        }
    }
}
