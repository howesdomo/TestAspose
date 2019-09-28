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

namespace XiaJBXie
{
    /// <summary>
    /// FrmTestFormula.xaml 的交互逻辑
    /// </summary>
    public partial class FrmTestFormula : Window
    {
        public FrmTestFormula()
        {
            InitializeComponent();
            this.btnTest.Click += BtnTest_Click;
        }

        private void BtnTest_Click(object sender, RoutedEventArgs e)
        {
            string error = Util.Excel.ExcelUtils_Aspose.TestAsposeCellsHotPatch();
            MessageBox.Show(error);
        }


    }
}
