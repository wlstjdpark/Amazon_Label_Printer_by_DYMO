using Microsoft.Win32;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace LabelsPrint
{
    public class Setting
    {
        public enum ProcessingOption : byte
        {
            None = 0,
            WriteToPng = 1,
            WriteToLabel = 2,
            Print = 4,            
        }

        public string FileName { get; private set; }
        public ProcessingOption Option { get; private set; }

        public Setting(string fileName, ProcessingOption option)
        {
            this.FileName = fileName;
            this.Option = option;            
        }        
    }    

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            PrintersComboBox.ItemsSource = DYMO.Label.Framework.Framework.GetLabelWriterPrinters();
            PrintersComboBox.SelectedIndex = 0;
            UpdatedPrintersComboBoxSelection();            

            FileNameTextBox.IsReadOnly = true;            
            CircularProgressBar.Visibility = Visibility.Hidden;

            LabelManager.OnUpdateStatusText += OnUpdateStatusText;
            LabelManager.OnSuccess += OnSuccess;
            LabelManager.OnFail += OnFail;            
        }

        private void UpdateStatusText(string text, Color? color = null)
        {
            var item = (StatusBarItem)Status.Items[0];
            var textBlock = (TextBlock)item.Content;
            textBlock.Text = text;
            textBlock.FontWeight = FontWeights.Bold;
            textBlock.Foreground = new SolidColorBrush(color ?? Colors.Black);                                                                      
        }

        private Setting.ProcessingOption GetOption()
        {
            Setting.ProcessingOption option = Setting.ProcessingOption.None;

            option |= WriteToPngCheckBox.IsChecked == true ? Setting.ProcessingOption.WriteToPng : option;
            option |= WriteToLabelCheckBox.IsChecked == true ? Setting.ProcessingOption.WriteToLabel : option;
            option |= PrintCheckBox.IsChecked == true ? Setting.ProcessingOption.Print : option;            

            return option;
        }        

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "pdf files (*.pdf)|*.pdf";
            if (openFileDialog.ShowDialog() == true)
            {
                FileNameTextBox.Text = openFileDialog.FileName;
                Task.Factory.StartNew(LabelManager.Start, new Setting(FileNameTextBox.Text, GetOption()));
                OnProcessing();
            }           
        }

        private void OnUpdateStatusText(string text, Color color)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                UpdateStatusText(text, color);
            }));
        }

        private void OnProcessing()
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                UpdateStatusText("Processing Image....", Colors.RoyalBlue);
                this.IsEnabled = false;
                CircularProgressBar.Visibility = Visibility.Visible;
            }));
        }

        private void OnFail()
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                UpdateStatusText("Fail!!", Colors.Red);
                this.IsEnabled = true;
                CircularProgressBar.Visibility = Visibility.Hidden;
            }));            
        }

        private void OnSuccess()
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                UpdateStatusText("Finished!!", Colors.RoyalBlue);
                this.IsEnabled = true;
                CircularProgressBar.Visibility = Visibility.Hidden;
            }));                        
        }        

        private void PrintersComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdatedPrintersComboBoxSelection();
        }

        private void UpdatedPrintersComboBoxSelection()
        {            
            var printer = PrintersComboBox.SelectedItem as DYMO.Label.Framework.ILabelWriterPrinter;            
            OpenFileButton.IsEnabled = printer != null;
            UpdateStatusText(OpenFileButton.IsEnabled ? "Click the \"File Open\"" : "Printer is not selected..",
                OpenFileButton.IsEnabled ? Colors.RoyalBlue : Colors.Red);

            if (printer == null)
            {
                return;
            }

            LabelManager.SelectedPrinter = printer.Name;
        }
    }
}
