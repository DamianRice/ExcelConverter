using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExcelConverter.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelConverter
{
    public delegate void SetProcessHandler(int process);

    public partial class MainForm : Form
    {
        #region Flieds以及Properties

        private Dictionary<Excel.XlFileFormat, string> _comboBoxData = new Dictionary<Excel.XlFileFormat, string>();
        private KeyValuePair<Excel.XlFileFormat, string> _covertFormat;
        private BindingSource _bs = new BindingSource();
        private string _fName = string.Empty;
        private string _fPath = string.Empty;
        private ExcelHandler _excelHandler = new ExcelHandler();
        private bool resultFile = false;
        private Dictionary<string, bool> resultFolder;
        Info info = new Info();

        public string FName
        {
            get => _fName;
            set
            {
                _fName = value;
                this.pathTextBox.Text = value;
            }
        }

        public string FPath
        {
            get => _fPath;
            set
            {
                _fPath = value;
                this.pathTextBox.Text = value;
            }
        }

        public KeyValuePair<Excel.XlFileFormat, string> CovertFormat
        {
            get => _covertFormat;
            set => _covertFormat = value;
        }

        public ExcelHandler Hand
        {
            get => _excelHandler;
            set => _excelHandler = value;
        }

        #endregion

        public MainForm()
        {
            InitializeComponent();
            
            Hand.SetProcess += UpdateProcess;
            Hand.SetTotal += SetTotalProcess;
        }

        #region 控件事件

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.Icon = Resources.excel;
            _comboBoxData.Add(Excel.XlFileFormat.xlOpenXMLWorkbook, ".xlsx");
            _comboBoxData.Add(Excel.XlFileFormat.xlExcel9795, ".xls");
            _bs.DataSource = _comboBoxData;
            this.modeSelection.DataSource = _bs;
            this.modeSelection.DisplayMember = "Value";
            this.modeSelection.ValueMember = "Key";
            this.batchmode.Checked = true;
            this.openPathBtn.Click += OpenDirBtn;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hand.CleanGC();
        }

        private void batchmode_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton s = (RadioButton) sender;

            //todo 要动态反射
            switch (s.Checked)
            {
                case true:

                    #region 切换打开文件dialog的方式

                    this.openPathBtn.Click -= OpenFileBtn;
                    this.openPathBtn.Click += OpenDirBtn;

                    #endregion

                    #region 按键名称切换

                    this.openPathBtn.Text = @"选择文件夹";

                    #endregion

                    break;
                case false:

                    #region 切换打开文件dialog的方式

                    this.openPathBtn.Click -= OpenDirBtn;
                    this.openPathBtn.Click += OpenFileBtn;

                    #endregion

                    #region 按键名称切换

                    this.openPathBtn.Text = @"选择文件";

                    #endregion

                    break;
            }
        }

        private void modeSelection_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox obj = (ComboBox) sender;
            CovertFormat = (KeyValuePair<Excel.XlFileFormat, string>) obj.SelectedItem;
        }

        private void startBtn_Click(object sender, EventArgs e)
        {
            Thread workerThread = new Thread(new ThreadStart(Start));
            workerThread.Start();
        }

        #endregion


        #region 切换委托

        private void OpenDirBtn(object sender, EventArgs e)
        {
            #region 选择路径

            FolderBrowserDialog folder = new FolderBrowserDialog();
            folder.Description = @"选择一个文件夹，将转换该文件夹下所有的Excel";
            folder.SelectedPath = System.Environment.CurrentDirectory;
            if (FPath != "")
            {
                folder.SelectedPath = FPath;
            }

            if (folder.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录  
                FPath = folder.SelectedPath;
            }

            #endregion
        }

        private void OpenFileBtn(object sender, EventArgs e)
        {
            #region 选择文件

            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.RestoreDirectory = true;
            fileDialog.Multiselect = false;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "Excel Files 2010 (.xlsx)|*.xlsx|" +
                                "Excel Files 2003 (.xls)|*.xls";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                FName = fileDialog.FileNames[0];
            }

            #endregion
        }

        private bool StartFileBtn()
        {
            if (!string.IsNullOrEmpty(FName) && File.Exists(FName))
            {
                var res = Hand.ConvertExcel(FName, CovertFormat.Key);
                return res;
            }
            else
            {
                MessageBox.Show(@"缺少路径,请选择文件或者文件夹");
                return false;
            }
        }

        private Dictionary<string, bool> StartFolderBtn()
        {
            if (!string.IsNullOrEmpty(FPath) && Directory.Exists(FPath))
            {
                var res = Hand.ConvertDirection(FPath, CovertFormat.Key);
                return res;
            }
            else
            {
                MessageBox.Show(@"缺少路径,请选择文件或者文件夹");
                return new Dictionary<string, bool> {{FPath, false}};
            }
        }

        #endregion

        #region 执行线程

        private void Start()
        {
            switch (batchmode.Checked)
            {
                case false:
                    resultFile = StartFileBtn();
                    this.info.GenerateInfo(FName,resultFile);
                    this.info.ShowDialog();
                    break;
                case true:
                    resultFolder = StartFolderBtn();
                    this.info.GenerateInfo(resultFolder, resultFolder.All(item=>item.Value));
                    this.info.ShowDialog();
                    break;
            }
        }

        #endregion

        #region 给子线程的委托，用来更新控件内容

        private void UpdateProcess(int process)
        {
            if (this.progressBar1.InvokeRequired)
            {
                this.BeginInvoke(new SetProcessHandler(UpdateProcess), new object[] {process});
            }
            else
            {
                progressBar1.Value = process;
            }
        }

        private void SetTotalProcess(int total)
        {
            if (this.progressBar1.InvokeRequired)
            {
                this.BeginInvoke(new SetProcessHandler(SetTotalProcess), new object[] {total});
            }
            else
            {
                this.progressBar1.Maximum = total;
            }
        }

        #endregion
    }
}