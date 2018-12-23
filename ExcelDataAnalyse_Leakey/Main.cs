using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CCWin;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelDataAnalyse_Leakey
{
    public partial class Main : Skin_Mac
    {
        //所有表格中的所有人
        List<ExamModel> allTesters = new List<ExamModel>();
        //被选出的人
        List<ExamModel> selectedTesters = new List<ExamModel>();
        //Excel文件地址
        List<string> ExcelFile = new List<string>();
        bool hasFilePath = false;
        public Main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            Start_btn.Enabled = false;
        }

        private void GetFile_btn_Click(object sender, EventArgs e)
        {
            selectPath();
        }

        private void selectPath()
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();   //显示选择文件对话框 
            openFileDialog1.Filter = "Excel 文件 |*.xlsx;*.xls";
            //openFileDialog1.Filter = "Excel 2003 文件 (*.xls)|*.xls";
            openFileDialog1.Multiselect = true;
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                    ExcelFile = new List<string>();
                    int fileCount = 0;
                    foreach (string fileName in openFileDialog1.FileNames)
                    {
                        fileCount++;
                        ExcelFile.Add(fileName);
                    }
                    this.filePathLBL.Text = "已选择：" + fileCount + "个文件";
                    hasFilePath = true;
                startBtnCheck();
            }
        }

        private void startBtnCheck()
        {
            if (hasFilePath)
            {
                Start_btn.Enabled = true;
            }
        }

        private void Start_btn_Click(object sender, EventArgs e)
        {
            //运行getData()方法，该方法返回一个string值，若运行正常，则自定义返回"success"
            if (getData().Equals("success"))
            {
                analyseData();
                showData();
            }
        }

        //获取数据
        private string getData()
        {
            if (ExcelFile == null)
            {
                MessageBox.Show("请重新选择文件~", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return "";
            }
            //用于存储数据的临时模型
            List<ExamModel> _tempExamModelList = new List<ExamModel>();
            //对于每一个Excel表格
            foreach (string fileName in ExcelFile)
            {
                IWorkbook workbook = null;  //新建IWorkbook对象 
                //姓名所在列
                int nameColumn = -1;
                //时间所在列
                int timeColumn = -1;
                //分数所在列
                int scoreColumn = -1;
                //标题所在行
                int titleRow = -1;

                FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
                {
                    try
                    {
                        workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("出现错误，请重新选择文件，错误内容："+e.ToString(),"提示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    }
                }
                else if (fileName.IndexOf(".xls") > 0) // 2003版本  
                {
                    try
                    {
                        workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("出现错误，请重新选择文件，错误内容：" + e.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                //找表头
                ISheet sheet1 = workbook.GetSheetAt(0);
                //对于sheet1的每一行
                for(int i = 0; i <= sheet1.LastRowNum; i++)
                {
                    IRow row = sheet1.GetRow(i);
                    if(row != null)
                    {
                        //对于该行的每一列
                        for(int j = 0; j <= row.LastCellNum; j++)
                        {
                            ICell cell = row.GetCell(j);
                            if(cell != null)
                            {
                                if(cell.ToString().Trim().Length != 0)
                                {
                                    switch (cell.ToString().Trim())
                                    {
                                        case "姓名":
                                            nameColumn = j;
                                            titleRow = i;
                                            break;
                                        case "分数":
                                            scoreColumn = j;
                                            titleRow = i;
                                            break;
                                        case "用时":
                                            timeColumn = j;
                                            titleRow = i;
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    //避免有人在姓名栏上填上“姓名”“分数”“和时间”关键字 设定一个跳出条件
                    if (nameColumn == -1 || scoreColumn == -1 || timeColumn == -1)
                    {
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
                if(nameColumn == -1 || scoreColumn== -1 || timeColumn == -1)
                {
                    MessageBox.Show("文件的表头不具备“姓名”，“分数”与“用时”中的任意一种，请检查。" , "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return "";
                }
                if(titleRow == -1)
                {
                    //根本就没找到标题
                    MessageBox.Show("未找到标题", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return "";
                }
                //找数据
                for(int k = titleRow +1; k <= sheet1.LastRowNum; k++)
                {
                    IRow row = sheet1.GetRow(k);
                    ExamModel _em = new ExamModel();
                    //找每一行的姓名
                    ICell nameCell = row.GetCell(nameColumn);
                    if (nameCell != null)
                    {
                        if(nameCell.ToString().Trim().Length != 0)
                        {
                            _em.fullName = nameCell.ToString().Trim();
                        }
                    }
                    //每一行的用时
                    ICell timeCell = row.GetCell(timeColumn);
                    if (timeCell != null)
                    {
                        if (timeCell.ToString().Trim().Length != 0)
                        {
                            _em.usingTime = timeCell.ToString().Trim();
                        }
                    }
                    //每一行的分数
                    ICell scoreCell = row.GetCell(scoreColumn);
                    if (scoreCell != null)
                    {
                        if (scoreCell.ToString().Trim().Length != 0)
                        {
                            _em.score = scoreCell.ToString().Trim();
                        }
                    }
                    _tempExamModelList.Add(_em);
                }
            }
            //赋值
            allTesters = _tempExamModelList;
            return "success";
        }

        //分析数据
        private void analyseData()
        {
            //对于获取到的每一个人
            int searchCount = 0;
            foreach (ExamModel _em in allTesters)
            {
                //首先确认筛选过的模型内没有该人名
                bool hasSameOne = false;
                //对于每一个已经被筛选过的人
                foreach(ExamModel _selectedEM in selectedTesters)
                {
                    //如果里面出现了此次循环的人名的话，跳过
                    if (_selectedEM.fullName.Equals(_em.fullName))
                    {
                        hasSameOne = true;
                        break;
                    }
                }
                //有的话跳过
                if (hasSameOne)
                {
                    continue;
                }
                //克隆一份 创建临时变量，用于加减
                ExamModel _tempEM = (ExamModel)_em.Clone();
                //此人的总分
                int score = 0;
                //把用时提出来
                string time = _tempEM.usingTime;
                //把分数栏转化为int
                int.TryParse(_tempEM.score.Split('分')[0], out score);
                //把重复的人名相关内容加起来，存入筛选过的模型中(equals为A字符串与B相等的时候)
                //foreach(ExamModel _compareEM in allTesters)
                for(int n = searchCount + 1 ;n< allTesters.Count; n++)
                {
                    //只找该人后面(searchCount+1或者更后面)的人
                    ExamModel _compareEM = allTesters[n];
                    if (_compareEM.fullName.Equals(_tempEM.fullName))
                    {
                        int compareScore = 0;
                        int.TryParse(_compareEM.score.Split('分')[0], out compareScore);
                        score = score + compareScore;
                        //时间提取出来，参与计算
                        string compareTime = _compareEM.usingTime;
                        //计算后返回的值给time变量
                        time = timeCalculate(time, compareTime);
                    }
                }
                //计算完成，存储
                _tempEM.scoreInt = score;
                _tempEM.score = score.ToString() + "分";
                //存储为 分钟-秒 的格式 便于排序
                if (time.Contains("分") && time.Contains("秒"))
                {
                    //分为秒数>10和秒数<10
                    int _second = -1;
                    int.TryParse(time.Split('分')[1].Split('秒')[0], out _second);
                    if (_second >= 10)
                    {
                        _tempEM.usingTime = time.Split('分')[0] + "-" + time.Split('分')[1].Split('秒')[0];
                    }
                    else if(_second >= 0)
                    {
                        _tempEM.usingTime = time.Split('分')[0] + "-0" + time.Split('分')[1].Split('秒')[0];
                    }

                }
                else if (time.Contains("分"))
                {
                    _tempEM.usingTime = time.Split('分')[0] + "-" + "00";
                }
                else if (time.Contains("秒"))
                {
                    _tempEM.usingTime = "00" + "-" + time.Split('秒')[0];
                }
                selectedTesters.Add(_tempEM);
                //避免找到自己，只往下找
                searchCount++;
            }

            //排序
            selectedTesters.Sort(delegate (ExamModel x, ExamModel y)
            {
                //比较分数
                if(y.scoreInt != x.scoreInt)
                {
                    return y.scoreInt.CompareTo(x.scoreInt);
                }
                else
                {//分数一样 比较用时
                    int xTime = 0;
                    int yTime = 0;
                    int.TryParse(x.usingTime.Replace("-",""), out xTime);
                    int.TryParse(y.usingTime.Replace("-",""), out yTime);
                    return xTime.CompareTo(yTime);
                }
            });


        }

        //把数据还是用excel展示
        private void showData()
        {
            //创建Excel文件名称
            FileStream fs = File.Create(Application.StartupPath + "\\对比结果.xls");
            //创建工作薄
            IWorkbook workbook = new HSSFWorkbook();

            //表格样式(粗体)
            ICellStyle boldStyle = workbook.CreateCellStyle();
            boldStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            boldStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            boldStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            boldStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            boldStyle.WrapText = true;
            boldStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//垂直
            boldStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            HSSFFont fontBold = (HSSFFont)workbook.CreateFont();
            fontBold.FontName = "宋体";//字体  
            fontBold.FontHeightInPoints = 10;//字号  
            fontBold.IsBold = true;//加粗  
            boldStyle.SetFont(fontBold);

            //(正常体)
            ICellStyle normalStyle = workbook.CreateCellStyle();
            normalStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.WrapText = true;
            normalStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            normalStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            normalStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//垂直
            normalStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");

            HSSFFont fontNormal = (HSSFFont)workbook.CreateFont();
            fontNormal.FontName = "宋体";//字体  
            fontNormal.FontHeightInPoints = 10;//字号  
            normalStyle.SetFont(fontNormal);

            //创建sheet
            ISheet sheet = workbook.CreateSheet("对比结果");
            //创建总人数+1行数的表格
            for (int i = 0; i < selectedTesters.Count + 1; i++)
            {
                IRow row = sheet.CreateRow(i);
                //创建标题
                if (i == 0)
                {
                    row.CreateCell(0).SetCellValue("排名");
                    row.GetCell(0).CellStyle = normalStyle;
                    row.CreateCell(1).SetCellValue("姓名");
                    row.GetCell(1).CellStyle = normalStyle;
                    row.CreateCell(2).SetCellValue("分数");
                    row.GetCell(2).CellStyle = normalStyle;
                    row.CreateCell(3).SetCellValue("用时");
                    row.GetCell(3).CellStyle = normalStyle;
                    continue;
                }
                //写数据
                else
                {
                    //名次写i的值
                    row.CreateCell(0).SetCellValue(i);
                    row.GetCell(0).CellStyle = boldStyle;
                    //姓名
                    row.CreateCell(1).SetCellValue(selectedTesters[i-1].fullName);
                    row.GetCell(1).CellStyle = normalStyle;
                    //分数
                    row.CreateCell(2).SetCellValue(selectedTesters[i - 1].scoreInt +"分");
                    row.GetCell(2).CellStyle = normalStyle;
                    //用时(化为x分x秒的格式)
                    row.CreateCell(3).SetCellValue(selectedTesters[i - 1].usingTime.Split('-')[0] + "分" + selectedTesters[i - 1].usingTime.Split('-')[1] + "秒");
                    row.GetCell(3).CellStyle = normalStyle;
                }
                
            }

            //向excel文件中写入数据并保保存
            workbook.Write(fs);
            fs.Close();
            System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo();
            info.FileName = Application.StartupPath + "\\对比结果.xls";
            try
            {
                System.Diagnostics.Process.Start(info);
            }
            catch (System.ComponentModel.Win32Exception we)
            {
                MessageBox.Show(this, we.Message);
                return;
            }
        }

        //这个方法用来计算用时
        private string timeCalculate(string originalTime, string addedTime)
        {
            //秒和分
            int originalSecond = 0;
            int originalMinute = 0;

            int addedSecond = 0;
            int addedMinute = 0;

            //取分钟
            //(contains表示字符串中含有某字符, tryparse是尝试转换为int型, split是从某个字符分开, [0]表示分开的第0部分（最左边）, trim表示去空格)
            if (originalTime.Contains("分"))
            {
                int.TryParse(originalTime.Split('分')[0].Trim(), out originalMinute);
            }
            if (addedTime.Contains("分"))
            {
                int.TryParse(addedTime.Split('分')[0].Trim(), out addedMinute);
            }
            //取秒数
            if (originalTime.Contains("秒"))
            {
                if(originalTime.Contains("分"))
                {
                    int.TryParse(originalTime.Split('秒')[0].Trim().Split('分')[1].Trim(), out originalSecond);
                }
                else
                {
                    int.TryParse(originalTime.Split('秒')[0].Trim(), out originalSecond);
                }
            }
            if (addedTime.Contains("秒"))
            {
                if (addedTime.Contains("分"))
                {
                    int.TryParse(addedTime.Split('秒')[0].Trim().Split('分')[1].Trim(), out addedSecond);
                }
                else
                {
                    int.TryParse(addedTime.Split('秒')[0].Trim(), out addedSecond);
                }
            }

            //计算
            originalSecond = addedSecond + originalSecond;
            originalMinute = addedMinute + originalMinute;
            //秒数>=60 分钟进1
            if(originalSecond >= 60)
            {
                originalMinute++;
                originalSecond = originalSecond - 60;
            }
            //返回源格式
            string time = originalMinute + "分" + originalSecond+"秒";
            return time;
        }
    }

    public class ExamModel : ICloneable
    {
        //需要排序和克隆↑
        //定义模型(预留了其他内容没有做)
        //账号
        public string account { get; set; }
        //昵称
        public string nickName { get; set; }
        //组织名
        public string organizationName { get; set; }
        //手机号
        public string phoneNumber { get; set; }
        //姓名
        public string fullName { get; set; }
        //单位
        public string unitName { get; set; }
        //职务
        public string dutyName { get; set; }
        //分数
        public string score { get; set; }
        //分数（int），为了排序方便
        public int scoreInt { get; set; }
        //用时
        public string usingTime { get; set; }

        //初始化
        public ExamModel()
        {
            account = "";
            nickName = "";
            organizationName = "";
            phoneNumber = "";
            fullName = "";
            unitName = "";
            dutyName = "";
            score = "";
            scoreInt = -1;
            usingTime = "";
        }

        //克隆方法
        public object Clone()
        {
            ExamModel _em = new ExamModel();
            _em.account = account;
            _em.nickName = nickName;
            _em.organizationName = organizationName;
            _em.phoneNumber = phoneNumber;
            _em.fullName = fullName;
            _em.unitName = unitName;
            _em.dutyName = dutyName;
            _em.score = score;
            _em.scoreInt = scoreInt;
            _em.usingTime = usingTime;

            return _em as object;//深复制
        }


    }
}
