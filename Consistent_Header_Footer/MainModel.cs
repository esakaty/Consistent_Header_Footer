/*
 * たぶんモデルに相当する処理
 */
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Threading;
using System.Windows;

namespace Consistent_Header_Footer
{
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

        public async void Execute(object? parameter)
        {
            _view.Bind_visibilityStatusBar = Visibility.Visible;
            _view.Bind_EnableOpelate = false;
            _view.Bind_txtStatusBar = "処理中";
            switch (parameter)
            {
                case "btnOpenFolderDialog":
                    break;

                case "btnOpen":
                    await Operate_Open();
                    break;

                case "btnChek":
                    await Opelate_Check();
                    break;

                case "btnStart":
                    await Opelate_Start();
                    break;

                default:
                    break;
            }
            _view.Bind_EnableOpelate = true;
            _view.Bind_visibilityStatusBar = Visibility.Hidden;
        }

        /// <summary>
        /// フォルダ展開
        /// 対象フォルダを開き、すべてのファイルをFileDataCollectionへ展開する。
        /// </summary>
        /// <returns></returns>
        private async System.Threading.Tasks.Task Operate_Open()
        {
            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    ObservableCollection<FileData> tmp_fileDatas = new();

                    var files = Directory.GetFiles(_view.PathFolder);
                    var fileNames = files.Select(file => System.IO.Path.GetFileName(file));
                    foreach (var (fileName, index) in fileNames.Select((filename, index) => (filename, index)))
                    {
                        Dispatcher dispatcher = Dispatcher.CurrentDispatcher;
                        dispatcher.Invoke((Action)(() =>
                        {
                            tmp_fileDatas.Add(new FileData()
                            {
                                Path = fileName,
                            });
                        }));
                    }
                    _view.Bind_FileDataCollection = tmp_fileDatas;
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        /// <summary>
        /// フォルダチェック
        /// </summary>
        /// <returns></returns>
        private async System.Threading.Tasks.Task Opelate_Check()
        {
            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    int count = 0;
                    int total = _view.Bind_FileDataCollection.Count() * 3;
                    foreach (var fileData in _view.Bind_FileDataCollection)
                    {
                        System.IO.FileInfo fi = new System.IO.FileInfo(fileData.Path);
                        if(fi.Extension.Substring(1, 3) == "doc")
                        {
                            Debug.Print("GetPageNum(Word):" + fileData.Path);
                            fileData.TotalPage = GetPageNum_Word(fileData);
                            fileData.UpdataView();
                        }
                        _view.Bind_valueStatusBar = (count++) * 100 / total;
                    }
                    foreach (var fileData in _view.Bind_FileDataCollection)
                    {
                        Debug.Print("GetGroup:" + fileData.Path);
                        GetGroup(fileData);
                        fileData.UpdataView();
                    }
                    foreach (var fileData in _view.Bind_FileDataCollection)
                    {
                        Debug.Print("GroupTotalPage:" + fileData.Path);
                        fileData.GroupTotalPage = _view.groupList[fileData.Group].GroupTotalPage;
                        fileData.UpdataView();
                        _view.Bind_valueStatusBar = (count++) * 100 / total;
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private async System.Threading.Tasks.Task Opelate_Start()
        {
            try
            {
                await System.Threading.Tasks.Task.Run(() =>
                {
                    foreach (FileData file in _view.Bind_FileDataCollection)
                    {
                        UpdateFooter_Word(file);
                    }
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
        }

        private int GetPageNum_Word(FileData file)
        {
            int pagenum = -1;
            if (!string.IsNullOrEmpty(file.Path))
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Open(_view.PathFolder + "\\" + file.Path);

                pagenum = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
                doc.Close();
                // todo 何かを待たないと最後のファイルがエラーになる。
                System.Threading.Thread.Sleep(500);
                wordApp.Quit();
            }
            return pagenum;
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
