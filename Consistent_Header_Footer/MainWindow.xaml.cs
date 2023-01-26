using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Printing;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
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
using System.Windows.Threading;
using Microsoft.Office.Interop.Word;
using static System.Collections.Specialized.BitVector32;

namespace Consistent_Header_Footer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
    }
    public enum eState
    {
        Wait,
        Operate,
        Stop,
        Complete,
        Skip,
        Error
    }
    public class FileData
    {
        public String Path { get; set; } = "";
        public int Group { get; set; }
        public int No { get; set; }
        public int StartPage { get; set; }
        public int TotalPage { get; set; }
        public int GroupTotalPage { get; set; }
        public eState State { get; set; }
        public bool Selected { get; set; }
    }
    public class GroupData
    {
        public string Name = "";
        public String Key = "";
        public int Count = 0;
        public int GroupTotalPage = 0;
    }

    ///============================================================
    /// ViewModel
    ///============================================================
    public class Main_ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        public ObservableCollection<FileData> FileDatas { get; set; }
        public List<GroupData> groupList { get; set; } = new List<GroupData>();
        public String PathFolder { get; set; }

        public String ValuePathFolder
        {
            get
            {
                return PathFolder;
            }
            set
            {
                PathFolder = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ValuePathFolder)));
            }
        }
        public ObservableCollection<FileData> ValueFileDataCollection
        {
            get
            {
                return FileDatas;
            }
            set
            {
                FileDatas = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ValueFileDataCollection)));
            }
        }

        public Command_Buttons Command_Buttons { get;  set; }

        public Main_ViewModel()
        {
            Command_Buttons = new Command_Buttons(this);
            FileDatas = new ObservableCollection<FileData>();
            PathFolder = System.IO.Path.GetDirectoryName(                System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\TargetFolder";
        }
    }

    public class Command_Buttons : ICommand
    {
        private Main_ViewModel _view { get; set; }
        public Command_Buttons(Main_ViewModel view) { _view = view; }

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object? parameter) { return true; }

        public void Execute(object? parameter)
        {
            switch (parameter)
            {
                case "btnOpen":
                    OpenFolder();
                    break;

                case "btnChek":
                    CheckFolder();
                    break;

                case "btnStart":
                    StartFolder();
                    break;

                default:
                    break;
            }
        }


        private void OpenFolder()
        {
            try
            {
                ObservableCollection<FileData> tmp_fileDatas = new();

                var files = Directory.GetFiles(_view.PathFolder);
                var fileNames = files.Select(file => System.IO.Path.GetFileName(file));
                foreach (var (fileName, index) in fileNames.Select((filename, index) => (filename, index)))
                {
                    tmp_fileDatas.Add(new FileData()
                    {
                        Path = fileName,
                    });
                }
                _view.ValueFileDataCollection = tmp_fileDatas;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private void CheckFolder()
        {
            foreach (var fileData in _view.ValueFileDataCollection)
            {
                GetPageNum(fileData);
            }
            foreach (var fileData in _view.ValueFileDataCollection)
            {
                GetGroup(fileData);
            }
            foreach (var data in _view.ValueFileDataCollection)
            {
                data.GroupTotalPage = _view.groupList[data.Group].GroupTotalPage;
            }
        }

        private void StartFolder()
        {
            foreach (FileData file in _view.ValueFileDataCollection)
            {
                UpdateFooter_Word(file);
            }
        }

        private void GetPageNum(FileData file)
        {
            if (!string.IsNullOrEmpty(file.Path))
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Open(_view.PathFolder + "\\" + file.Path);

                file.TotalPage = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
                doc.Close();
                // todo 何かを待たないと最後のファイルがエラーになる。
                System.Threading.Thread.Sleep(500);
                wordApp.Quit();
            }
        }

        private void GetGroup(FileData file)
        {
            bool flg = false;
            int tmpPageNum = 1;
            string Group_Name = "";

            foreach (var (Group, index) in _view.groupList.Select((GroupData, Index) => (GroupData, Index)))
            {
                if (Group_Name == Group.Name)
                {
                    flg = true;
                    tmpPageNum = Group.GroupTotalPage + 1;

                    file.Group = index;
                    file.No = Group.Count + 1;
                    Group.Count++;
                    Group.GroupTotalPage += file.TotalPage;
                    break;
                }
            }
            if (flg != true)
            {
                _view.groupList.Add(new GroupData()
                {
                    Name = Group_Name,
                    GroupTotalPage = file.TotalPage,
                    Key = ""
                });
            }
            file.StartPage = tmpPageNum;
        }

        private void UpdateFooter_Word(FileData file)
        {
            Microsoft.Office.Interop.Word.Application? wordApp = null;
            Microsoft.Office.Interop.Word.Document? doc = null;

            file.State = eState.Operate;
            try
            {
                System.Diagnostics.Debug.WriteLine("Start\n Path=" + file.Path);

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
                file.State = eState.Complete;

                // todo 何かを待たないと最後のファイルがエラーになる。
                System.Threading.Thread.Sleep(300);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
                file.State = eState.Error;
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
    }
}
