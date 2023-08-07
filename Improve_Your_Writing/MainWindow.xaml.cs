using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Improve_Your_Writing_Core;
using Microsoft.Win32;
using NPOI.SS.Formula.Functions;

namespace Improve_Your_Writing
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public DocumentSettings DocumentSettingsInstance { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            DocumentSettingsInstance = new DocumentSettings()
            {
                // 初始化属性值
                FontSize = 24,
                FontName = "Arial",
                OutputDocxPath = "output.docx",
                InputXlsxPath = "input.xlsx",
                StartAfterLine = 0
            };
            DataContext = DocumentSettingsInstance;
        }

        

        private void Button_ChooseXlsx_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Xlsx Document (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                TextBox_XlsxPath.Text = openFileDialog.FileName;
            }
        }

        private void Button_ChooseDocx_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Word Document (*.docx)|*.docx";
            if (saveFileDialog.ShowDialog() == true)
            {
                TextBox_DocxPath.Text = saveFileDialog.FileName;
            }
        }

        private void OK_Click(object sender, RoutedEventArgs e)
        {
            Core.Run(DocumentSettingsInstance);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            TextBox_FontSize.Text = (1 + int.Parse(TextBox_FontSize.Text)).ToString();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            TextBox_FontSize.Text = (Math.Max(int.Parse(TextBox_FontSize.Text) - 1, 0)).ToString();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            TextBox_StartAfterLine.Text = (1 + int.Parse(TextBox_StartAfterLine.Text)).ToString();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            TextBox_StartAfterLine.Text = (Math.Max(int.Parse(TextBox_StartAfterLine.Text) - 1, 0)).ToString();
        }
    }
}
