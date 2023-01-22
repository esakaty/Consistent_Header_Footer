using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Printing;
using System.Text;
using System.Text.RegularExpressions;
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

        public class FileData
        {
            public enum eState
            {
                Wait,
                Operate,
                Stop,
                Complete,
                Skip,
                Error
            }

            public String Path ="";
            public int Group = 0;
            public int No = 0;
            public int StartPage = 0;
            public int TotalPage = 0;
            public int GroupTotalPage = 0;
            public eState State = eState.Skip;
            public bool Selected = false;
        }

        public class GroupData
        {
            public string Name = "";
            public String Key = "";
            public int Count = 0;
        }

        public List<FileData> fileDatas { get; set; } = new List<FileData>();
        public List<GroupData> groupList { get; set; } = new List<GroupData>();

        public MainWindow()
        {
            InitializeComponent();
            this.txtPathFolder.Text = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\TargetFolder";
        }

        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            Update_All_Files();
        }

        private void Update_All_Files()
        {
            foreach (FileData file in fileDatas)
            {
                UpdateFooter_Word(file);
            }
        }

        private void UpdateFooter_Word(FileData file)
        {
            Microsoft.Office.Interop.Word.Application? wordApp = null;
            Microsoft.Office.Interop.Word.Document? doc = null;

            file.State = FileData.eState.Operate;
            try
            {
                System.Diagnostics.Debug.WriteLine("Start\n Path="+file.Path);

                // Wordアプリケーションを起動
                wordApp = new Microsoft.Office.Interop.Word.Application();
                System.Diagnostics.Debug.WriteLine(" Run Word app");

                // 新しい文書を開く
                doc = wordApp.Documents.Open(file.Path);
                System.Diagnostics.Debug.WriteLine(" Open Word file");

                // 全セクションに対して
                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {
                    if (section.Index == 1)
                    {
                        //開始番号を変更
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].PageNumbers.RestartNumberingAtSection = true;
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].PageNumbers.StartingNumber = file.StartPage;
                    }

                    // フッターを取得
                    var primaryFooter = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    // フッターにページ番号フィールドを埋め込む
                    primaryFooter.Range.Fields.Add(primaryFooter.Range, WdFieldType.wdFieldPage);
                    primaryFooter.Range.InsertBefore("( ");
                    primaryFooter.Range.InsertAfter(" / " + file.GroupTotalPage + " )");
                    primaryFooter.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
                System.Diagnostics.Debug.WriteLine(" Add Footer");

                // 更新した文書を保存
                doc.Save();
                System.Diagnostics.Debug.WriteLine(" Save File");
                file.State = FileData.eState.Complete;

                // todo 何かを待たないと最後のファイルがエラーになる。
                System.Threading.Thread.Sleep(300);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
                file.State = FileData.eState.Error;
            }
            finally
            {
                if (doc != null)
                {
                    System.Diagnostics.Debug.WriteLine(" Close Word file");
                    // ファイルクローズ
                    doc.Close();
                }
                if (wordApp != null)
                {
                    System.Diagnostics.Debug.WriteLine(" Quit Word app");
                    // Wordを終了
                    wordApp.Quit();
                }
            }
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<FileData> tmp_fileDatas = new List<FileData>();

                var files = Directory.GetFiles(this.txtPathFolder.Text);
                var fileNames = files.Select(file => System.IO.Path.GetFullPath(file));
                foreach (var (fileName, index) in fileNames.Select((filename, index) => (filename, index)))
                {
                    tmp_fileDatas.Add(new FileData()
                    {
                        Path = fileName,
                        No = index,
                        State = FileData.eState.Wait
                    });
                    UpdateFileData(tmp_fileDatas[index]);
                }
                fileDatas = tmp_fileDatas;

                foreach(var data in fileDatas)
                {
                    bool flg = false;
                    int tmpPageNum = 1;
                    string Group_Name = "";

                    foreach (var (Group,index) in groupList.Select((GroupData,Index) => (GroupData,Index)))
                    {
                        if(Group_Name == Group.Name)
                        {
                            flg = true;
                            tmpPageNum = Group.Count + 1;

                            data.Group = index;
                            Group.Count+=data.TotalPage;
                            break;
                        }
                    }
                    if(flg != true)
                    {
                        groupList.Add(new GroupData()
                        {
                            Name = Group_Name,
                            Count = data.TotalPage,
                            Key = ""
                        });
                    }
                    data.StartPage = tmpPageNum;
                }

                foreach (var data in fileDatas)
                {
                    data.GroupTotalPage = groupList[data.Group].Count;
                }

                System.Data.DataTable dt = new System.Data.DataTable();
                dt = new System.Data.DataTable();
                dt.Columns.Add("FileName");
                dt.Columns.Add("Group");
                dt.Columns.Add("No");
                dt.Columns.Add("StartPage");
                dt.Columns.Add("TotalPage");
                dt.Columns.Add("GroupTotalPage");
                dt.Columns.Add("State");
                dt.Columns.Add("Select");

                for (int i = 0; i < fileDatas.Count; i++)
                {
                    System.Data.DataRow row = dt.NewRow();
                    row[0] = System.IO.Path.GetFileName(fileDatas[i].Path);
                    row[1] = fileDatas[i].Group;
                    row[2] = fileDatas[i].No;
                    row[3] = fileDatas[i].StartPage;
                    row[4] = fileDatas[i].TotalPage;
                    row[5] = groupList[fileDatas[i].Group].Count;
                    row[6] = fileDatas[i].State;
                    row[7] = fileDatas[i].Selected;
                    dt.Rows.Add(row);
                }

                dataGrid_FileInfo.DataContext = dt;
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private void UpdateFileData(FileData file)
        {
            if (!string.IsNullOrEmpty(file.Path))
            {
                var folderPath = file.Path;
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var pageCounts = new Dictionary<string, int>();
                var doc = wordApp.Documents.Open(file.Path);
                file.TotalPage = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
                doc.Close();
                // todo 何かを待たないと最後のファイルがエラーになる。
                System.Threading.Thread.Sleep(500);
                wordApp.Quit();
            }
        }
    }
}
