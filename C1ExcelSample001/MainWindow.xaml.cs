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
using C1.WPF.Excel;

namespace C1ExcelSample001
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // C1XLBookがExcelブックを表すクラス
            C1XLBook book = new C1XLBook();

            // XLSheetがExcelのシートを表す
            XLSheet sheet = book.Sheets[0];

            // 1行目の1から3列目に値を設定する
            sheet[0, 0].Value = 1;
            sheet[0, 1].Value = 2;
            sheet[0, 2].Value = 3;

            // スタイルを設定する
            XLStyle style = new XLStyle(book);
            style.ForeColor = Colors.Blue;

            sheet[0, 1].Style = style;

            // 式を用いる
            // 1行目の1から3列のSUM(合計)を4列目に求める
            sheet[0, 3].Formula = "SUM(A1: C1)";

            // 画像を設定する
            WriteableBitmap img = new WriteableBitmap(new BitmapImage(new Uri("icon.png", UriKind.Relative)));
            sheet[1, 0].Value = img;


            // 保存する
            book.Save(@"C:\<ドキュメントフォルダのパス>\mybook.xls");
        }
    }
}
