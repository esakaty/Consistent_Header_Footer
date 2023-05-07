using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Printing;
using System.Reflection;
using System.Runtime.CompilerServices;
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
using Microsoft.VisualBasic;
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
    public class FileData : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;
        private void RaisePropertyChanged([CallerMemberName] string propertyName = "") => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private String _Path {  get; set; } = "";
        public String Path {    get=> _Path;    set{_Path = value;RaisePropertyChanged();}}

        private int _Group {    get; set; }
        public int Group {      get=>_Group;    set{_Group = value;RaisePropertyChanged();}}

        private int _No {       get; set; }
        public int No {         get=>_No;       set{_No = value;RaisePropertyChanged();}}

        private int _StartPage{ get; set; }
        public int StartPage{   get=> _StartPage;set{_StartPage = value;RaisePropertyChanged();}}

        private int _TotalPage{ get; set; }
        public int TotalPage{   get=> _TotalPage;set{_TotalPage = value;RaisePropertyChanged();}}

        private int _GroupTotalPage{    get; set; }
        public int GroupTotalPage{      get => _GroupTotalPage; set { _GroupTotalPage = value; RaisePropertyChanged(); } }

        private eState _State { get; set; }
        public eState State { get => _State; set { _State = value; RaisePropertyChanged(); } }

        private bool _Selected { get; set; }
        public bool Selected { get => _Selected; set { _Selected = value; RaisePropertyChanged(); } }

        public void UpdataView()
        {
            DispatcherFrame frame = new DispatcherFrame();
            var callback = new DispatcherOperationCallback(obj =>
            {
                ((DispatcherFrame)obj).Continue = false;
                return null;
            });
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, callback, frame);
            Dispatcher.PushFrame(frame);
        }
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
        public Main_ViewModel()
        {
            Command_Buttons = new Command_Buttons(this);
            FileDatas = new ObservableCollection<FileData>();
            PathFolder = System.IO.Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\TargetFolder";

            bEnableOpelate = true;

            visibilityStatusBar = Visibility.Hidden;
            valueStatusBar = 0;
            txtStatusBar = "処理中";
        }
        public void UpdataView()
        {
            DispatcherFrame frame = new DispatcherFrame();
            var callback = new DispatcherOperationCallback(obj =>
            {
                ((DispatcherFrame)obj).Continue = false;
                return null;
            });
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, callback, frame);
            Dispatcher.PushFrame(frame);
        }

        public Command_Buttons Command_Buttons { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
        public ObservableCollection<FileData> FileDatas { get; set; }
        public int FileDataIndex;
        public List<GroupData> groupList { get; set; } = new List<GroupData>();
        public String PathFolder { get; set; }
        public bool bEnableOpelate { get; set; }
        public String txtStatusBar { get; set; }
        public Visibility visibilityStatusBar { get; set; }
        public int valueStatusBar { get; set; }


        public ObservableCollection<FileData> Bind_FileDataCollection
        {
            get
            {
                return FileDatas;
            }
            set
            {
                FileDatas = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bind_FileDataCollection)));
                UpdataView();
            }
        }

        public String Bind_PathFolder
        {
            get
            {
                return PathFolder;
            }
            set
            {
                PathFolder = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bind_PathFolder)));
                UpdataView();
            }
        }
        public bool Bind_EnableOpelate {
            get
            {
                return bEnableOpelate;
            }
            set
            {
                bEnableOpelate = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bind_EnableOpelate)));
                UpdataView();
            }
        }
        public String Bind_txtStatusBar
        {
            get
            {
                return txtStatusBar;
            }
            set
            {
                txtStatusBar = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bind_txtStatusBar)));
                UpdataView();
            }
        }

        public Visibility Bind_visibilityStatusBar
        {
            get
            {
                return visibilityStatusBar;
            }
            set
            {
                visibilityStatusBar = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bind_visibilityStatusBar)));
                UpdataView();
            }
        }
        public int Bind_valueStatusBar
        {
            get
            {
                return valueStatusBar;
            }
            set
            {
                valueStatusBar = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Bind_valueStatusBar)));
                UpdataView();
            }
        }
    }
}
