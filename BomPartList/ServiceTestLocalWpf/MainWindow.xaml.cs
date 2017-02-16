namespace ServiceTestLocalWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            MwGridUc.Children.Add(new TestServiceUc());
        }
    }
}
