using EncooTech.PdfConversionActivity.Helper;
using EncooTech.PdfConversionActivity.Model;
using EncooTech.PdfConversionActivity.Services;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace EncooTech.PdfConversionActivity.Designer
{
    // PdfConversionActivityDesigner.xaml 的交互逻辑
    public partial class PdfConversionDialogDesigner
    {
        public string DialogTargetFormat { get; set; }

        public string DialogResultPath { get; set; }

        private ObservableCollection<PdfPathsList> pdfPathsList;
        public PdfConversionDialogDesigner()
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
            Um.NavigateUri = new Uri(@"https://marketplace.encoo.com/#/activity/detail?lang=zh-cn&packageId=EncooTech.PdfConversionActivity");
            List<string> TargetFormats = new List<string> { "Word", "Excel", "图片", "TXT" };
            CbTargetFormat.ItemsSource = TargetFormats;
            this.Closing += PdfConversionDialogDesigner_Closing;
        }

        private void PdfConversionDialogDesigner_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            DialogResult = true;
        }

        public PdfConversionDialogDesigner(List<string> files, string targetFormat) : this()
        {
            pdfPathsList = new ObservableCollection<PdfPathsList>();

            files?.ForEach(c =>
            {
                setInfo(c);
            });
            list.ItemsSource = pdfPathsList;
            list.Items.Refresh();
            if (!string.IsNullOrEmpty(targetFormat))
            {
                CbTargetFormat.SelectedValue = targetFormat;
            }
            else
            {
                CbTargetFormat.SelectedIndex = 0;
            }
        }

        private void setInfo(string c)
        {
            var relativePath = FileHelper.GetFileFullPath(c);
            DirectoryInfo info = new DirectoryInfo(relativePath);
            DialogResultPath = info.Parent.FullName;
            pdfPathsList.Add(new PdfPathsList(c));
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            //设置girid的选中元素为Button所在行的元素
            list.SelectedItem = ((System.Windows.Controls.Button)sender).DataContext;
            //在数据集合中删除此元素                       
            pdfPathsList.RemoveAt(list.SelectedIndex);
            list.Items.Refresh();//刷新listview
        }

        private void BtnChooseFile_Click(object sender, RoutedEventArgs e)
        {
            var fileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PDF文件|*.pdf",
                Multiselect = true
            };
            var close = fileDialog.ShowDialog();
            if (close != null && close.Value)
            {
                double  FileMaxWidch = 0;
                pdfPathsList = new ObservableCollection<PdfPathsList>();
                var fileNames = fileDialog.FileNames;
                foreach (var fileName in fileNames)
                {
                    FileMaxWidch = fileName.Length> FileMaxWidch ? fileName.Length: FileMaxWidch;
                    setInfo(fileName);
                }
                list.ItemsSource = pdfPathsList;
                  
                //使listview根据内容自动调整宽度
                if (list.View is GridView gv)
                {
                    foreach (GridViewColumn gvc in gv.Columns)
                    {
                        if (gvc.Header.ToString() == "文件列表")
                        {
                            if (FileMaxWidch < 45)
                            {
                                gvc.Width = 100;
                                gvc.Width = Double.NaN;
                            }
                            else
                            {
                                gvc.Width = 520;
                            }
                        }                      
                    }
                }
                list.Items.Refresh();
            }
        }

        private void StartConversion_Click(object sender, RoutedEventArgs e)
        {
            if (pdfPathsList == null || pdfPathsList.Count == 0)
            {
                return;
            }
            string Type = Convert.ToString(CbTargetFormat.SelectedValue);
            string TargetFilePath = "";
            PdfConversionToolsServices services = new PdfConversionToolsServices();
            List<string> iamgeNameList;
            List<Task> taskList = new List<Task>();
            switch (Type)
            {
                case "Word":
                    foreach (var c in pdfPathsList)
                    {
                        taskList.Add(Task.Factory.StartNew(() =>
                        {
                            services.ConvertWordFromPdf(c.PdfPath, TargetFilePath, true, true, "DOCX");
                        }));
                    };
                    break;
                case "图片":
                    foreach (var c in pdfPathsList)
                    {
                        taskList.Add(Task.Factory.StartNew(() =>
                        {
                            services.ConvertImageFromPdf(c.PdfPath, TargetFilePath, "png", System.IO.Path.GetFileNameWithoutExtension(c.pdfPath), 0, out iamgeNameList);
                        }));
                    };
                    break;
                case "Excel":
                    foreach (var c in pdfPathsList)
                    {
                        taskList.Add(Task.Factory.StartNew(() =>
                        {
                            services.ConvertExcelFromPdf(c.PdfPath, TargetFilePath, true);
                        }));
                    };
                    break;
                case "TXT":
                    foreach (var c in pdfPathsList)
                    {
                        taskList.Add(Task.Factory.StartNew(() =>
                        {
                            services.ConvertTextFromPdf(c.PdfPath, TargetFilePath);
                        }));
                    };
                    break;
                default: break;
            }
            try
            {
                int index = Task.WaitAny(taskList.ToArray());
            }
            catch (AggregateException ae)
            {

                foreach (var ex in ae.InnerExceptions)
                {
                    throw new Exception(ex.Message);
                }
            }
            MessageBoxEx.Show(this, "         转换完成", "提示");
        }

        private void BrowseResults_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(DialogResultPath))
            {
                Process proc = new Process();
                proc.StartInfo.FileName = "explorer";
                //打开资源管理器
                proc.StartInfo.Arguments = DialogResultPath;
                proc.Start();
            }
        }

        private void Window_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Um_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink link = sender as Hyperlink;
            Process.Start(new ProcessStartInfo(link.NavigateUri.AbsoluteUri));
        }

        private void RPA_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink link = sender as Hyperlink;
            Process.Start(new ProcessStartInfo(link.NavigateUri.AbsoluteUri));
        }
    }
}
