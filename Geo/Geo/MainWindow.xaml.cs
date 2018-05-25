using Microsoft.Win32;
using System;
using System.Threading;
using System.Windows;
using System.Windows.Input;

namespace Geo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private HttpGetHelper httpGetHelper = new HttpGetHelper();

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Location.Text = httpGetHelper.GaoDeAnalysis("key=3e0ded4b2852e194c63565d151c2e606&address=" + Address.Text);
        }

        private static bool isBusy = false;
        private static int targetColumnInt;
        private static int originalColumnInt;
        private static int rowsInt;
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //ProcessExcel.ProcessSchoolCode();
            //return;

            if (targetColumn.Text == "" || originalColumn.Text == "" || rows.Text == "")
            {
                MessageBox.Show("请输入地址所在的列和经纬度要写入的列，还有需要转换的总行数。");
                return;
            }

            OpenFileDialog dialog = new OpenFileDialog
            {
                Multiselect = false,// 该值确定是否可以选择多个文件
                Title = "请选择Excel文件",
                Filter = "Excel文件(*.xlsx)|*.xlsx"
            };

            if (dialog.ShowDialog() == true)
            {
                if (!isBusy)
                {
                    isBusy = true;
                    targetColumnInt = int.Parse(targetColumn.Text);
                    originalColumnInt = int.Parse(originalColumn.Text);
                    rowsInt = int.Parse(rows.Text);
                    Thread ProcessData = new Thread(() =>
                    {
                        ProcessExcel.Process(dialog.FileName, originalColumnInt, targetColumnInt, rowsInt);
                        isBusy = false;
                    });
                    ProcessData.Start();
                }
                else
                {
                    MessageBox.Show("正在批量处理中...");
                }
            }
        }

        #region 鼠标移动窗口
        /// <summary>
        /// 鼠标移动窗口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainTitle_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
        #endregion 鼠标移动窗口

        #region 最小化按钮
        /// <summary>
        /// 最小化按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnMin_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }
        #endregion 最小化按钮
        
        #region 关闭程序按钮
        /// <summary>
        /// 关闭程序按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        #endregion 关闭程序按钮

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("         第十组 WXH");
        }
    }
}
