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

using Microsoft.Office.Interop.Word;
using static System.Collections.Specialized.BitVector32;

namespace Consistent_Header_Footer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        enum e_OperateResult{
            Error,
            Complete,
            Continue
        };

        public MainWindow()
        {
            InitializeComponent();
            this.txtPathFolder.Text = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\TargetFolder";
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            e_OperateResult result = e_OperateResult.Continue;
            result = UpdateFooter_Word(this.txtPathFolder.Text + "\\TargetFile1.docx",2,10);
        }

        private e_OperateResult UpdateFooter_Word(string strFilePath,int startPageNum,int totalPageNum)
        {
            e_OperateResult result = e_OperateResult.Error;
            if (System.IO.File.Exists(this.txtPathFolder.Text + "\\TargetFile1.docx")) {
                // Wordアプリケーションを起動
                var wordApp = new Microsoft.Office.Interop.Word.Application();

                // 新しい文書を開く
                var doc = wordApp.Documents.Open(strFilePath);

                // 全セクションに対して
                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {   
                    if (section.Index == 0)
                    {
                        //開始番号を変更
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].PageNumbers.RestartNumberingAtSection = true;
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].PageNumbers.StartingNumber = 2;
                    }

                    // フッターを取得
                    var primaryFooter = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    // フッターにページ番号フィールドを埋め込む
                    primaryFooter.Range.Fields.Add(primaryFooter.Range, WdFieldType.wdFieldPage);
                    primaryFooter.Range.InsertBefore("( ");
                    primaryFooter.Range.InsertAfter(" / " + totalPageNum + " )");

                }

                // 更新した文書を保存
                doc.Save();

                // Wordを終了
                wordApp.Quit();
            }

            return result;
        }
    }

}
