using System.Windows.Controls;

namespace Word_Graf.Elements
{
    /// <summary>
    /// Логика взаимодействия для Owner.xaml
    /// </summary>
    public partial class Owner : UserControl
    {
        public Owner(Context.OwnerContext roomOwner)
        {
            InitializeComponent();
            NameOwner.Content = $"{roomOwner.LastName} {roomOwner.FirstName} {roomOwner.SurName}";
        }
    }
}
