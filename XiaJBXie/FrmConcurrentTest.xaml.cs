using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace XiaJBXie
{
    /// <summary>
    /// FrmConcurrentTest.xaml 的交互逻辑
    /// </summary>
    public partial class FrmConcurrentTest : Window
    {
        FrmConcurrentTestViewModel ViewModel { get; set; }

        TestMSExcelServiceReference.PCWebServiceSoapClient Web { get; set; }

        public FrmConcurrentTest()
        {
            InitializeComponent();
            this.Web = new TestMSExcelServiceReference.PCWebServiceSoapClient();
            this.Web.Endpoint.Address = new System.ServiceModel.EndpointAddress(new Uri("http://192.168.1.215:27891/TestMSExcel/PCWebService.asmx"));

            this.ViewModel = new FrmConcurrentTestViewModel();
            this.DataContext = this.ViewModel;
            this.initEvent();
        }

        private void initEvent()
        {
            this.btnRun.Click += BtnRun_Click;
            this.btnStop.Click += BtnStop_Click;
        }

        BackgroundWorker bgWorker { get; set; }

        private void BtnRun_Click(object sender, RoutedEventArgs e)
        {
            if (bgWorker != null && bgWorker.IsBusy == true)
            {
                MessageBox.Show("正在运行");
                return;
            }

            if (bgWorker == null)
            {
                bgWorker = new BackgroundWorker();
                bgWorker.WorkerReportsProgress = true;
                bgWorker.DoWork += BgWorker_DoWork;
                bgWorker.ProgressChanged += BgWorker_ProgressChanged;
                bgWorker.RunWorkerCompleted += BgWorker_RunWorkerCompleted;
            }

            bgWorker.RunWorkerAsync();
        }

        private bool mContinueRun { get; set; }

        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string arg = "测试{0}".FormatWith(this.ViewModel.Number);
            mContinueRun = true;
            while (mContinueRun)
            {
                try
                {
                    string r = this.Web.TestConcurrentRead(arg);
                    if (r.Equals(arg))
                    {
                        bgWorker.ReportProgress(1);
                    }
                    else
                    {
                        bgWorker.ReportProgress(0);
                    }
                }
                catch (Exception)
                {
                    bgWorker.ReportProgress(-1);
                }

                System.Threading.Thread.Sleep(100);
            }
        }

        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.ViewModel.RunCount = this.ViewModel.RunCount + 1;

            switch (e.ProgressPercentage)
            {
                case 0:
                    this.ViewModel.FailCount = this.ViewModel.FailCount + 1;
                    break;

                case 1:
                    ;
                    break;

                case -1:
                default:
                    this.ViewModel.ErrorCount = this.ViewModel.ErrorCount + 1;
                    break;
            }
        }

        private void BgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.GetFullInfo());
            }
            else
            {
                MessageBox.Show("停止执行");

                this.ViewModel.RunCount = 0;
                this.ViewModel.FailCount = 0;
                this.ViewModel.ErrorCount = 0;
            }
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            mContinueRun = false;
        }

    }

    public class FrmConcurrentTestViewModel : BaseViewModel
    {

        public string _Number;
        public string Number
        {
            get { return _Number; }
            set
            {
                _Number = value;
                this.OnPropertyChanged("Number");
            }
        }

        private int _RunCount;

        public int RunCount
        {
            get { return _RunCount; }
            set
            {
                _RunCount = value;
                this.OnPropertyChanged("RunInfo");
            }
        }

        private int _FailCount;

        public int FailCount
        {
            get { return _FailCount; }
            set
            {
                _FailCount = value;
                this.OnPropertyChanged("FailCount");
            }
        }



        private int _ErrorCount;

        public int ErrorCount
        {
            get { return _ErrorCount; }
            set
            {
                _ErrorCount = value;
                this.OnPropertyChanged("ErrorCount");
            }
        }

        public string RunInfo
        {
            get
            {
                string r = string.Empty;

                r = "运行次数 : {0}; 失败次数 : {1}; 报错次数 : {2}".FormatWith(this.RunCount, this.FailCount, this.ErrorCount);

                return r;
            }
        }
    }
}
