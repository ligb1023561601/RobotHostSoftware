using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.Windows.Forms.DataVisualization.Charting;
using System.IO;
using System.Data.OleDb;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.Util;

namespace Robot_HostSoft
{
    public partial class Form1 : Form
    {
        public SerialPort PortA = new SerialPort();
        public SerialPort PortB = new SerialPort();
        public bool isPortAExist = false;
        public bool isPortBExist = false;
        public bool isSetPropertyA=false;
        public bool isSetPropertyB=false;
        public Byte[] PortA_Rec_Buffer = new Byte[200];                         //A口接收缓冲
        public Byte[] PortB_Rec_Buffer = new Byte[200];                         //B口接收缓冲
        public int PortAReceivedCount = 0;                                      //串口收数存在缓冲区的起始索引
        public int PortBReceivedCount = 0;                                      //串口收数存在缓冲区的起始索引
        public List<int> X_Axis_PositionList = new List<int>();
        public List<int> Y_Axis_PositionList = new List<int>();
        public List<int> Z_Axis_PositionList = new List<int>();
        public List<int> F_Axis_PositionList = new List<int>();
        public List<int> TensionForceList = new List<int>();

        public List<DateTime> X_Axis_TimeList = new List<DateTime>();
        public List<DateTime> Y_Axis_TimeList = new List<DateTime>();
        public List<DateTime> Z_Axis_TimeList = new List<DateTime>();
        public List<DateTime> F_Axis_TimeList = new List<DateTime>();

        public int PortARecvState = 0;
        public int PortBRecvState = 0;
        public const int RECV_HEAD1 = 0;
        public const int RECV_HEAD2 = 1;
        public const int RECV_ORIGIN_ID = 2; 
        public const int RECV_DATA = 3;
        public const int RECV_CHECKSUM = 4;
        public const int RECV_END1 = 5;
        public const int RECV_END2 =6;
        public int DataSourceFlag = 0;                                          //0:XY轴；1：ZF轴
        //各轴的相对位移
        public int X_Axis_RelativeDisplacement = 0;                             
        public int Y_Axis_RelativeDisplacement = 0;
        public int Z_Axis_RelativeDisplacement = 0;
        public int F_Axis_RelativeDisplacement = 0;
        /// <summary>
        /// 窗口初始化的设置
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            this.MaximumSize = this.MaximumSize;
            this.MinimumSize = this.MinimumSize;
            
            //串口A设置
            for (int i = 0; i < 10; i++)
			{
                Cbx_Port_A.Items.Add("COM" + i);
			}
            cbxBaud_A.Items.Add("1200");
            cbxBaud_A.Items.Add("2400");
            cbxBaud_A.Items.Add("4800");
            cbxBaud_A.Items.Add("9600");
            cbxBaud_A.Items.Add("19200");
            cbxBaud_A.Items.Add("38400");
            cbxBaud_A.Items.Add("115200");
            cbxBaud_A.SelectedIndex = 3;

            cbxStopbit_A.Items.Add("0");
            cbxStopbit_A.Items.Add("1");
            cbxStopbit_A.Items.Add("1.5");
            cbxStopbit_A.Items.Add("2");
            cbxStopbit_A.SelectedIndex = 0;

            cbxData_A.Items.Add("8");
            cbxData_A.Items.Add("7");
            cbxData_A.Items.Add("6");
            cbxData_A.Items.Add("5");
            cbxData_A.SelectedIndex = 0;

            cbxOddbit_A.Items.Add("奇校验");
            cbxOddbit_A.Items.Add("偶校验");
            cbxOddbit_A.Items.Add("none");
            cbxOddbit_A.SelectedIndex = 2;

            //串口B设置
            for (int j = 0; j < 10; j++)
            {
                cbx_PortID_B.Items.Add("COM" + j);
            }
            cbxBaud_B.Items.Add("1200");
            cbxBaud_B.Items.Add("2400");
            cbxBaud_B.Items.Add("4800");
            cbxBaud_B.Items.Add("9600");
            cbxBaud_B.Items.Add("19200");
            cbxBaud_B.Items.Add("38400");
            cbxBaud_B.Items.Add("115200");
            cbxBaud_B.SelectedIndex = 3;

            cbx_Stopbit_B.Items.Add("0");
            cbx_Stopbit_B.Items.Add("1");
            cbx_Stopbit_B.Items.Add("1.5");
            cbx_Stopbit_B.Items.Add("2");
            cbx_Stopbit_B.SelectedIndex = 0;

            cbx_Data_B.Items.Add("8");
            cbx_Data_B.Items.Add("7");
            cbx_Data_B.Items.Add("6");
            cbx_Data_B.Items.Add("5");
            cbx_Data_B.SelectedIndex = 0;

            cbx_Oddbit_B.Items.Add("奇校验");
            cbx_Oddbit_B.Items.Add("偶校验");
            cbx_Oddbit_B.Items.Add("none");
            cbx_Oddbit_B.SelectedIndex = 2;

            //X轴chart设置
            chart_X.Titles.Add("X轴位置");
            chart_X.Series[0].ChartType = SeriesChartType.Line;
            chart_X.Series[0].BorderWidth = 2;
            chart_X.Series[0].Color = System.Drawing.Color.LightSeaGreen;
            chart_X.Series[0].LegendText = "X轴坐标";

            chart_X.ChartAreas[0].AxisX.Enabled = AxisEnabled.True;
            chart_X.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
            chart_X.ChartAreas[0].AxisX.Name = "时间";
            chart_X.ChartAreas[0].AxisY.Name = "坐标位置";
            chart_X.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart_X.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart_X.ChartAreas[0].AxisX.Title = "时间";
            chart_X.ChartAreas[0].AxisY.Title = "位置坐标";
            chart_X.ChartAreas[0].AxisX.MajorGrid.Interval = 0.001;
            chart_X.ChartAreas[0].AxisY.MajorGrid.Interval = 1000;
            chart_X.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            chart_X.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            chart_X.ChartAreas[0].CursorX.AutoScroll = true;
            chart_X.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart_X.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_X.ChartAreas[0].CursorX.LineColor = Color.Blue;
            chart_X.ChartAreas[0].CursorY.AutoScroll = true;
            chart_X.ChartAreas[0].CursorY.IsUserEnabled = true;
            chart_X.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_X.ChartAreas[0].CursorY.LineColor = Color.Blue;

            //Y轴chart设置
            chart_Y.Titles.Add("Y轴位置");
            chart_Y.Series[0].ChartType = SeriesChartType.Line;
            chart_Y.Series[0].BorderWidth = 2;
            chart_Y.Series[0].Color = System.Drawing.Color.LightSeaGreen;
            chart_Y.Series[0].LegendText = "Y轴坐标";

            chart_Y.ChartAreas[0].AxisX.Enabled = AxisEnabled.True;
            chart_Y.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
            chart_Y.ChartAreas[0].AxisX.Name = "时间";
            chart_Y.ChartAreas[0].AxisY.Name = "坐标位置";
            chart_Y.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart_Y.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart_Y.ChartAreas[0].AxisX.Title = "时间";
            chart_Y.ChartAreas[0].AxisY.Title = "位置坐标";
            chart_Y.ChartAreas[0].AxisX.MajorGrid.Interval = 10;
            chart_Y.ChartAreas[0].AxisY.MajorGrid.Interval = 1000;
            chart_Y.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            chart_Y.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            chart_Y.ChartAreas[0].CursorX.AutoScroll = true;
            chart_Y.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart_Y.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_Y.ChartAreas[0].CursorX.LineColor = Color.Blue;
            chart_Y.ChartAreas[0].CursorY.AutoScroll = true;
            chart_Y.ChartAreas[0].CursorY.IsUserEnabled = true;
            chart_Y.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_Y.ChartAreas[0].CursorY.LineColor = Color.Blue;

            //Z轴chart设置
            chart_Z.Titles.Add("Z轴位置");
            chart_Z.Series[0].ChartType = SeriesChartType.Line;
            chart_Z.Series[0].BorderWidth = 2;
            chart_Z.Series[0].Color = System.Drawing.Color.LightSeaGreen;
            chart_Z.Series[0].LegendText = "Z轴坐标";

            chart_Z.ChartAreas[0].AxisX.Enabled = AxisEnabled.True;
            chart_Z.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
            chart_Z.ChartAreas[0].AxisX.Name = "时间";
            chart_Z.ChartAreas[0].AxisY.Name = "坐标位置";
            chart_Z.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart_Z.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart_Z.ChartAreas[0].AxisX.Title = "时间";
            chart_Z.ChartAreas[0].AxisY.Title = "位置坐标";
            chart_Z.ChartAreas[0].AxisX.MajorGrid.Interval = 10;
            chart_Z.ChartAreas[0].AxisY.MajorGrid.Interval = 1000;
            chart_Z.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            chart_Z.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            chart_Z.ChartAreas[0].CursorX.AutoScroll = true;
            chart_Z.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart_Z.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_Z.ChartAreas[0].CursorX.LineColor = Color.Blue;
            chart_Z.ChartAreas[0].CursorY.AutoScroll = true;
            chart_Z.ChartAreas[0].CursorY.IsUserEnabled = true;
            chart_Z.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_Z.ChartAreas[0].CursorY.LineColor = Color.Blue;

            chart_F.Titles.Add("肌张力电机位置");
            chart_F.Series[0].ChartType = SeriesChartType.Line;
            chart_F.Series[0].BorderWidth = 2;
            chart_F.Series[0].Color = System.Drawing.Color.LightSeaGreen;
            chart_F.Series[0].LegendText = "肌张力电机坐标";

            chart_F.ChartAreas[0].AxisX.Enabled = AxisEnabled.True;
            chart_F.ChartAreas[0].AxisY.Enabled = AxisEnabled.True;
            chart_F.ChartAreas[0].AxisX.Name = "时间";
            chart_F.ChartAreas[0].AxisY.Name = "坐标位置";
            chart_F.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart_F.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart_F.ChartAreas[0].AxisX.Title = "时间";
            chart_F.ChartAreas[0].AxisY.Title = "位置坐标";
            chart_F.ChartAreas[0].AxisX.MajorGrid.Interval = 0.1;
            chart_F.ChartAreas[0].AxisY.MajorGrid.Interval = 1000;
            chart_F.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot;
            chart_F.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

            chart_F.ChartAreas[0].CursorX.AutoScroll = true;
            chart_F.ChartAreas[0].CursorX.IsUserEnabled = true;
            chart_F.ChartAreas[0].CursorX.IsUserSelectionEnabled = true;
            chart_F.ChartAreas[0].CursorX.LineColor = Color.Blue;
            chart_F.ChartAreas[0].CursorY.AutoScroll = true;
            chart_F.ChartAreas[0].CursorY.IsUserEnabled = true;
            chart_F.ChartAreas[0].CursorY.IsUserSelectionEnabled = true;
            chart_F.ChartAreas[0].CursorY.LineColor = Color.Blue;

            btnMotorEnable.Enabled = false;

            导出数据为ExcelToolStripMenuItem.Click += toolStripBtn_Export_Click;
            导入数据ToolStripMenuItem.Click += toolStripBtn_Import_Click;

            btnRecord_X.Click += btnRecord_X_Click;
            btnRecord_Y.Click += btnRecord_X_Click;
            btnRecord_Z.Click += btnRecord_X_Click;
            btnRecord_F.Click += btnRecord_X_Click;

            记录数据ToolStripMenuItem.Click += btnRecord_X_Click;

            this.skinEngine1.SkinFile = @"MP10.ssk";
        }

        /// <summary>
        /// 串口A默认设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SetPortDefault_A_Click(object sender, EventArgs e)
        {
            cbxBaud_A.SelectedIndex = 6;
            cbxData_A.SelectedIndex = 0;
            cbxStopbit_A.SelectedIndex = 1;
            cbxOddbit_A.SelectedIndex = 2;

        }

        /// <summary>
        /// 串口A检测
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPotrACheck_Click(object sender, EventArgs e)
        {
            this.CheckAvailavlablePort();
        }

        /// <summary>
        /// 串口B默认设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_SetPortDefault_B_Click(object sender, EventArgs e)
        {
            cbxBaud_B.SelectedIndex = 6;
            cbx_Data_B.SelectedIndex = 0;
            cbx_Stopbit_B.SelectedIndex = 1;
            cbx_Oddbit_B.SelectedIndex = 2;
        }

        /// <summary>
        /// 串口B存在检查
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPotrBCheck_Click(object sender, EventArgs e)
        {
            this.CheckAvailavlablePort();
        }

        /// <summary>
        /// 检查串口存在
        /// </summary>
        public void CheckAvailavlablePort()
        {
            this.isPortAExist = false;
            this.isPortBExist = false;
            Cbx_Port_A.Items.Clear();
            cbx_PortID_B.Items.Clear();
            try
            {
                string[] PortsName= SerialPort.GetPortNames();  
                if (PortsName.Length!=0)
                {
                    foreach (var item in PortsName)
                    {
                        Cbx_Port_A.Items.Add(item);                        
                        cbx_PortID_B.Items.Add(item);                        
                    }
                    Cbx_Port_A.SelectedIndex = 0;
                    cbx_PortID_B.SelectedIndex = 1;
                    this.isPortAExist = true;
                    this.isPortBExist = true;
                }
                else
                {
                    MessageBox.Show("串口不存在，请检查连接！", "Warning");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        /// <summary>
        ///检查串口设置 
        /// </summary>
        /// <param name="PortID"></param>
        /// <returns></returns>
        public bool CheckPortSetup(string PortID)
        {
            switch (PortID)
            {   
                case "A":
                    if (Cbx_Port_A.Text=="")
                    {
                        return false;
                    }
                    if (cbxBaud_A.Text=="")
                    {
                        return false;
                    }
                    if (cbxData_A.Text=="")
                    {
                        return false;
                    }
                    if (cbxStopbit_A.Text=="")
                    {
                        return false;
                    }
                    if (cbxOddbit_A.Text=="")
                    {
                        return false;
                    }
                    break;
                case "B":
                    if (cbx_PortID_B.Text == "")
                    {
                        return false;
                    }
                    if (cbxBaud_B.Text=="")
                    {
                        return false;
                    }
                    if (cbx_Data_B.Text=="")
                    {
                        return false;
                    }
                    if (cbx_Stopbit_B.Text=="")
                    {
                        return false;
                    }
                    if (cbx_Oddbit_B.Text=="")
                    {
                        return false;
                    }
                    break;
                default:
                    break;
            }
            return true;
        }
        /// <summary>
        /// 设置串口的属性
        /// </summary>
        public void SetupPortsProperty()
        {
            //ProA Property 
            PortA.PortName = Cbx_Port_A.SelectedItem.ToString();
            PortA.BaudRate =Convert.ToInt32(cbxBaud_A.Text);
            PortA.DataBits = Convert.ToInt32(cbxData_A.Text);
            switch (cbxStopbit_A.Text)
            {
                case "0":
                    PortA.StopBits = StopBits.None;
                    break;
                case "1":
                    PortA.StopBits = StopBits.One;
                    break;
                case "1.5":
                    PortA.StopBits = StopBits.OnePointFive;
                    break;
                case "2":
                    PortA.StopBits = StopBits.Two;
                    break;
                default:
                    break;
            }
            switch (cbxOddbit_A.Text)
            {
                case "奇校验":
                    PortA.Parity = Parity.Odd;
                    break;
                case "偶校验":
                    PortA.Parity = Parity.Even;
                    break;
                case "none":
                    PortA.Parity = Parity.None;
                    break;
                default:
                    break;
            }
            PortA.ReadTimeout = 1000;
            PortA.WriteTimeout = 1000;
            this.isSetPropertyA = true;
            //PortB Property
            PortB.PortName = cbx_PortID_B.Text;
            PortB.BaudRate = Convert.ToInt32(cbxBaud_B.Text);
            PortB.DataBits = Convert.ToInt32(cbx_Data_B.Text);
            switch (cbx_Stopbit_B.Text)
            {
                case "0":
                    PortB.StopBits = StopBits.None;
                    break;
                case "1":
                    PortB.StopBits = StopBits.One;
                    break;
                case "1.5":
                    PortB.StopBits = StopBits.OnePointFive;
                    break;
                case "2":
                    PortB.StopBits = StopBits.Two;
                    break;
                default:
                    break;
            }
            switch (cbx_Oddbit_B.Text)
            {
                case "奇校验":
                    PortB.Parity = Parity.Odd;
                    break;
                case "偶校验":
                    PortB.Parity = Parity.Even;
                    break;
                case "none":
                    PortB.Parity = Parity.None;
                    break;
                default:
                    break;
            }
            PortB.ReadTimeout = 1000;
            PortB.WriteTimeout = 1000;
            this.isSetPropertyB = true;

            PortA.DataReceived += new SerialDataReceivedEventHandler(PortADataRecieved);
            PortB.DataReceived += new SerialDataReceivedEventHandler(PortBDataRecieved);
        }

        /// <summary>
        /// 打开串口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPortOpen_Click(object sender, EventArgs e)
        {
            if (btnPortOpen.Text=="打开串口")
            {
                if ((!CheckPortSetup("A")) || (!CheckPortSetup("B")))
                {
                    MessageBox.Show("串口未设置好，请检查！", "Warning");
                    return;
                }

                if (!PortA.IsOpen)
                {
                    try
                    {
                        this.SetupPortsProperty();
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("串口A未设置好，请检查！", "Warning");
                    }
                    try
                    {
                        PortA.Open();
                        statuslbOpenPortNameA.Text = PortA.PortName;
                        statuslbPortPropertyA.Text = "波特率："+PortA.BaudRate.ToString() +" 数据位："+ PortA.DataBits.ToString() +
                            " 停止位:"+PortA.StopBits.ToString()+" 奇偶校验："+PortA.Parity.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("串口A" + ex.ToString());
                    }
                }

                if (!PortB.IsOpen)
                {
                    if (!isSetPropertyB)
                    {
                        MessageBox.Show("串口B未设置好，请检查！", "Warning");
                    }                                     
                    try
                    {
                        PortB.Open();
                        statuslbOpenPortNameB.Text = PortB.PortName;
                        statuslbPortPropertyB.Text = "波特率：" + PortB.BaudRate.ToString() + " 数据位：" + PortB.DataBits.ToString() +
                            " 停止位:" + PortB.StopBits.ToString() + " 奇偶校验：" + PortB.Parity.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("串口B" + ex.ToString());
                    }
                }
                if (PortA.PortName.Equals(PortB.PortName))
                {
                    MessageBox.Show("串口名称重复，请重新设置！", "Warning");
                    return;
                }

                else
                {
                    btnPortOpen.Text = "关闭串口";
                    Cbx_Port_A.Enabled = false;
                    cbxBaud_A.Enabled = false;
                    cbxData_A.Enabled = false;
                    cbxStopbit_A.Enabled = false;
                    cbxOddbit_A.Enabled = false;
                    btn_SetPortDefault_A.Enabled = false;
                    btnPotrACheck.Enabled = false;

                    cbx_PortID_B.Enabled = false;
                    cbxBaud_B.Enabled = false;
                    cbx_Data_B.Enabled = false;
                    cbx_Stopbit_B.Enabled = false;
                    cbx_Oddbit_B.Enabled = false;
                    btn_SetPortDefault_B.Enabled = false;
                    btnPotrBCheck.Enabled = false;

                    btnMotorEnable.Enabled = true;
                    导入数据ToolStripMenuItem.Enabled = false;                  //打开串口后不允许导入数据
                    toolStripBtn_Import.Enabled = false;
                }
            }
            else
            {
                try
                {
                    PortA.Close();
                    PortB.Close();
                    this.isSetPropertyA = false;
                    this.isSetPropertyB = false;
                    btnPortOpen.Text = "打开串口";
                    Cbx_Port_A.Enabled = true;
                    cbxBaud_A.Enabled = true;
                    cbxData_A.Enabled = true;
                    cbxStopbit_A.Enabled = true;
                    cbxOddbit_A.Enabled = true;
                    btn_SetPortDefault_A.Enabled = true;
                    btnPotrACheck.Enabled = true;

                    cbx_PortID_B.Enabled = true;
                    cbxBaud_B.Enabled = true;
                    cbx_Data_B.Enabled = true;
                    cbx_Stopbit_B.Enabled = true;
                    cbx_Oddbit_B.Enabled = true;
                    btn_SetPortDefault_B.Enabled = true;
                    btnPotrBCheck.Enabled = true;

                    btnMotorEnable.Enabled = false;
                    BtnEnable(false);

                    导入数据ToolStripMenuItem.Enabled = true;                  //打开串口后不允许导入数据
                    toolStripBtn_Import.Enabled = true;
                }
                catch (Exception)
                {
                    MessageBox.Show("串口关闭失败！", "Warning");
                }
            }
            
           

        }
        /// <summary>
        /// 串口A接收事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PortADataRecieved(object sender, SerialDataReceivedEventArgs e)
        {
            this.Invoke((EventHandler)(delegate{
                int ReceivedDataLength_A = 0;                                 //串口一次收数长度
                Byte RecvCheckSum_A = 0;                                       //校验和

                int PortA_Data_Length = 0;                                    //包中数据字节长度
                Byte[] PortA_Data_Buffer = new Byte[10];
                Byte[] X_Axis_Buf = new Byte[2];
                Byte[] Y_Axis_Buf = new Byte[2];
                Byte[] Z_Axis_Buf = new Byte[2];
                Byte[] F_Axis_Buf = new Byte[2];
                Byte[] Force_Sensor_Buf = new Byte[2];
                Byte PortAByteBuffer = new Byte();                            //每次判断时用的字节缓存
                Byte[] ReceivedData = null;
                try
                {
                    if (PortA.BytesToRead > 0)
                    {
                        ReceivedData = new Byte[PortA.BytesToRead];
                        ReceivedDataLength_A = PortA.Read(ReceivedData, 0, PortA.BytesToRead);
                    }
                    Array.Copy(ReceivedData, 0, PortA_Rec_Buffer, PortAReceivedCount, ReceivedData.Length);
                    PortAReceivedCount += ReceivedDataLength_A;
                    if (PortAReceivedCount > 15)
                    {
                        for (int i = 0; i < PortAReceivedCount; i++)
                        {
                            PortAByteBuffer = PortA_Rec_Buffer[i];
                            switch (PortARecvState)
                            {
                                case RECV_HEAD1:
                                    if (PortAByteBuffer == 0x66)
                                    {
                                        PortARecvState = RECV_HEAD2;
                                        RecvCheckSum_A = PortAByteBuffer;
                                    }
                                    break;
                                case RECV_HEAD2:
                                    if (PortAByteBuffer == 0x77)
                                    {
                                        PortARecvState = RECV_ORIGIN_ID;
                                        RecvCheckSum_A += PortAByteBuffer;
                                    }
                                    else
                                    {
                                        PortARecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_ORIGIN_ID:
                                    if (PortAByteBuffer == 0x0A)
                                    {
                                        PortARecvState = RECV_DATA;
                                        RecvCheckSum_A += PortAByteBuffer;
                                        DataSourceFlag = 0;
                                    }
                                    else if (PortAByteBuffer == 0x0B)
                                    {
                                        PortARecvState = RECV_DATA;
                                        RecvCheckSum_A += PortAByteBuffer;
                                        DataSourceFlag = 1;
                                    }
                                    else
                                    {
                                        PortARecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_DATA:
                                    PortA_Data_Buffer[PortA_Data_Length] = PortAByteBuffer;
                                    RecvCheckSum_A += PortAByteBuffer;
                                    PortA_Data_Length++;
                                    if (PortA_Data_Length > 9)
                                    {
                                        PortARecvState = RECV_CHECKSUM;
                                        PortA_Data_Length = 0;
                                    }
                                    break;
                                case RECV_CHECKSUM:
                                    if (RecvCheckSum_A == PortAByteBuffer)
                                    {
                                        PortARecvState = RECV_END1;
                                    }
                                    else
                                    {
                                        PortARecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_END1:
                                    if (PortAByteBuffer == 0x88)
                                    {
                                        PortARecvState = RECV_END2;
                                    }
                                    else
                                    {
                                        PortARecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_END2:                                 //一个包完全无误之后
                                    if (PortAByteBuffer == 0x99)
                                    {
                                        PortAReceivedCount = 0;                 //无误则下次从缓冲区的0索引处放入
                                        //XY轴数据
                                        if (DataSourceFlag == 0)
                                        {
                                            X_Axis_Buf[1] = PortA_Data_Buffer[0];                //1对应高位，0对应低位
                                            X_Axis_Buf[0] = PortA_Data_Buffer[1];
                                            X_Axis_PositionList.Add(BitConverter.ToInt16(X_Axis_Buf, 0));
                                            X_Axis_TimeList.Add(DateTime.Now);
                                            chart_X.Series[0].XValueType = ChartValueType.DateTime;   //不指定这格式会出错
                                            chart_X.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_X.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_X.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_X.Series[0].Points.DataBindXY(X_Axis_TimeList, X_Axis_PositionList);
                                            chart_X.DataBind();
                                            
                                            Y_Axis_Buf[1] = PortA_Data_Buffer[2];
                                            Y_Axis_Buf[0] = PortA_Data_Buffer[3];
                                            Y_Axis_PositionList.Add(BitConverter.ToInt16(Y_Axis_Buf, 0));
                                            Y_Axis_TimeList.Add(DateTime.Now);
                                            chart_Y.Series[0].XValueType = ChartValueType.DateTime;   //不指定这格式会出错
                                            chart_Y.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_Y.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_Y.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_Y.Series[0].Points.DataBindXY(Y_Axis_TimeList, Y_Axis_PositionList);
                                            chart_Y.DataBind();
                                            while (X_Axis_PositionList.Count>500)
                                            {
                                                RemoveData();                                                 
                                            }
                                        }
                                        //ZF轴以及力数据
                                        if (DataSourceFlag == 1)
                                        {
                                            Z_Axis_Buf[1] = PortA_Data_Buffer[4];
                                            Z_Axis_Buf[0] = PortA_Data_Buffer[5];
                                            Z_Axis_PositionList.Add(BitConverter.ToInt16(Z_Axis_Buf, 0));
                                            Z_Axis_TimeList.Add(DateTime.Now);
                                            chart_Z.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_Z.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_Z.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_Z.Series[0].Points.DataBindXY(Z_Axis_TimeList, Z_Axis_PositionList);
                                            chart_Z.DataBind();

                                            F_Axis_Buf[1] = PortA_Data_Buffer[6];
                                            F_Axis_Buf[0] = PortA_Data_Buffer[7];
                                            F_Axis_PositionList.Add(BitConverter.ToInt16(F_Axis_Buf, 0));
                                            F_Axis_TimeList.Add(DateTime.Now);
                                            chart_F.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_F.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_F.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_F.Series[0].XValueType = ChartValueType.DateTime;
                                            chart_F.Series[0].Points.DataBindXY(F_Axis_TimeList, F_Axis_PositionList);
                                            chart_F.DataBind();

                                            Force_Sensor_Buf[1] = PortA_Data_Buffer[8];
                                            Force_Sensor_Buf[0] = PortA_Data_Buffer[9];
                                            TensionForceList.Add(BitConverter.ToInt16(Force_Sensor_Buf, 0));
                                            while (Z_Axis_PositionList.Count>500)
                                            {
                                                RemoveData();
                                            }
                                        }
                                    }
                                    PortARecvState = RECV_HEAD1;
                                    PortA.DiscardInBuffer();
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    PortA.DiscardInBuffer();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }));
           
        }
        /// <summary>
        /// 串口B接收事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PortBDataRecieved(object sender, SerialDataReceivedEventArgs e)
        {
            this.Invoke((EventHandler)(delegate
            {
                int ReceivedDataLength_B = 0;                                 //串口一次收数长度
                Byte RecvCheckSum_B = 0;                                       //校验和

                int PortB_Data_Length = 0;                                    //包中数据字节长度
                Byte[] PortB_Data_Buffer = new Byte[10];
                Byte[] X_Axis_Buf = new Byte[2];
                Byte[] Y_Axis_Buf = new Byte[2];
                Byte[] Z_Axis_Buf = new Byte[2];
                Byte[] F_Axis_Buf = new Byte[2];
                Byte[] Force_Sensor_Buf = new Byte[2];
                Byte PortBByteBuffer = new Byte();                            //每次判断时用的字节缓存
                Byte[] ReceivedData = null;
                try
                {
                    if (PortB.BytesToRead > 0)
                    {
                        ReceivedData = new Byte[PortB.BytesToRead];
                        ReceivedDataLength_B = PortB.Read(ReceivedData, 0, PortB.BytesToRead);
                    }
                    Array.Copy(ReceivedData, 0, PortB_Rec_Buffer, PortBReceivedCount, ReceivedData.Length);
                    PortBReceivedCount += ReceivedDataLength_B;
                    if (PortBReceivedCount > 15)
                    {
                        for (int i = 0; i < PortBReceivedCount; i++)
                        {
                            PortBByteBuffer = PortB_Rec_Buffer[i];
                            switch (PortBRecvState)
                            {
                                case RECV_HEAD1:
                                    if (PortBByteBuffer == 0x66)
                                    {
                                        PortBRecvState = RECV_HEAD2;
                                        RecvCheckSum_B = PortBByteBuffer;
                                    }
                                    break;
                                case RECV_HEAD2:
                                    if (PortBByteBuffer == 0x77)
                                    {
                                        PortBRecvState = RECV_ORIGIN_ID;
                                        RecvCheckSum_B += PortBByteBuffer;
                                    }
                                    else
                                    {
                                        PortBRecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_ORIGIN_ID:
                                    if (PortBByteBuffer == 0x0A)
                                    {
                                        PortBRecvState = RECV_DATA;
                                        RecvCheckSum_B += PortBByteBuffer;
                                        DataSourceFlag = 0;
                                    }
                                    else if (PortBByteBuffer == 0x0B)
                                    {
                                        PortBRecvState = RECV_DATA;
                                        RecvCheckSum_B += PortBByteBuffer;
                                        DataSourceFlag = 1;
                                    }
                                    else
                                    {
                                        PortBRecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_DATA:
                                    PortB_Data_Buffer[PortB_Data_Length] = PortBByteBuffer;
                                    RecvCheckSum_B += PortBByteBuffer;
                                    PortB_Data_Length++;
                                    if (PortB_Data_Length > 9)
                                    {
                                        PortBRecvState = RECV_CHECKSUM;
                                        PortB_Data_Length = 0;
                                    }
                                    break;
                                case RECV_CHECKSUM:
                                    if (RecvCheckSum_B == PortBByteBuffer)
                                    {
                                        PortBRecvState = RECV_END1;
                                    }
                                    else
                                    {
                                        PortBRecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_END1:
                                    if (PortBByteBuffer == 0x88)
                                    {
                                        PortBRecvState = RECV_END2;
                                    }
                                    else
                                    {
                                        PortBRecvState = RECV_HEAD1;
                                    }
                                    break;
                                case RECV_END2:                                 //一个包完全无误之后
                                    if (PortBByteBuffer == 0x99)
                                    {
                                        PortBReceivedCount = 0;                 //无误则下次从缓冲区的0索引处放入
                                        //XY轴数据
                                        if (DataSourceFlag == 0)
                                        {
                                            X_Axis_Buf[1] = PortB_Data_Buffer[0];                //1对应高位，0对应低位
                                            X_Axis_Buf[0] = PortB_Data_Buffer[1];
                                            X_Axis_PositionList.Add(BitConverter.ToInt16(X_Axis_Buf, 0));
                                            X_Axis_TimeList.Add(DateTime.Now);
                                            chart_X.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_X.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_X.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_X.Series[0].XValueType = ChartValueType.DateTime;
                                            chart_X.Series[0].Points.DataBindXY(X_Axis_TimeList, X_Axis_PositionList);
                                            chart_X.DataBind();                                            
                                            Y_Axis_Buf[1] = PortB_Data_Buffer[2];
                                            Y_Axis_Buf[0] = PortB_Data_Buffer[3];
                                            Y_Axis_PositionList.Add(BitConverter.ToInt16(Y_Axis_Buf, 0));
                                            Y_Axis_TimeList.Add(DateTime.Now);                                           
                                            chart_Y.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_Y.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_Y.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_Y.Series[0].XValueType = ChartValueType.DateTime;
                                            chart_Y.Series[0].Points.DataBindXY(Y_Axis_TimeList, Y_Axis_PositionList);
                                            chart_Y.DataBind();
                                            while (X_Axis_PositionList.Count>500)
                                            {
                                                RemoveData();
                                            }
                                        }
                                        //ZF轴以及力数据
                                        if (DataSourceFlag == 1)
                                        {
                                            Z_Axis_Buf[1] = PortB_Data_Buffer[4];
                                            Z_Axis_Buf[0] = PortB_Data_Buffer[5];
                                            Z_Axis_PositionList.Add(BitConverter.ToInt16(Z_Axis_Buf, 0));
                                            Z_Axis_TimeList.Add(DateTime.Now);
                                            chart_Z.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_Z.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_Z.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_Z.Series[0].XValueType = ChartValueType.DateTime;
                                            chart_Z.Series[0].Points.DataBindXY(Z_Axis_TimeList, Z_Axis_PositionList);
                                            chart_Z.DataBind();

                                            F_Axis_Buf[1] = PortB_Data_Buffer[6];
                                            F_Axis_Buf[0] = PortB_Data_Buffer[7];
                                            F_Axis_PositionList.Add(BitConverter.ToInt16(F_Axis_Buf, 0));
                                            F_Axis_TimeList.Add(DateTime.Now);
                                            chart_F.ChartAreas[0].AxisX.LabelStyle.Format = "HH:mm:ss";
                                            chart_F.ChartAreas[0].AxisX.MajorGrid.Enabled = true;
                                            chart_F.ChartAreas[0].AxisY.MajorGrid.Enabled = true;
                                            chart_F.Series[0].XValueType = ChartValueType.DateTime;
                                            chart_F.Series[0].Points.DataBindXY(F_Axis_TimeList, F_Axis_PositionList);
                                            chart_F.DataBind();
                                            Force_Sensor_Buf[1] = PortB_Data_Buffer[8];
                                            Force_Sensor_Buf[0] = PortB_Data_Buffer[9];
                                            TensionForceList.Add(BitConverter.ToInt16(Force_Sensor_Buf, 0));
                                            while (Z_Axis_PositionList.Count>500)
                                            {
                                                RemoveData();
                                            }
                                        }
                                    }
                                    PortBRecvState = RECV_HEAD1;
                                    PortB.DiscardInBuffer();
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                    PortB.DiscardInBuffer();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            })
            );
            
        }
        /// <summary>
        /// 电机使能
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMotorEnable_Click(object sender, EventArgs e)
        {
            try
            {
                switch (btnMotorEnable.Text)
                {
                    case "电机使能":
                        if (PortA.IsOpen == true)
                        {
                            if (PortB.IsOpen == true)
                            {
                                Byte[] command = new Byte[16];
                                command[0] = 0x66;
                                command[1] = 0x77;
                                command[2] = 0xFF;
                                command[3] = 0xFF;
                                command[14] = 0x88;
                                command[15] = 0x99;
                                PortA.Write(command, 0, command.Length);
                                PortB.Write(command, 0, command.Length);
                                btnMotorEnable.Text = "电机停机";
                                BtnEnable(true);
                                lb_XAxis.Text = "X轴位移：0mm";
                                lb_YAxis.Text = "Y轴位移：0mm";
                                lb_ZAxis.Text = "Z轴位移：0mm";
                                lb_FAxis.Text = "F轴位移：0mm";
                            }
                            else
                            {
                                MessageBox.Show("串口B未打开！", "Warning");
                            }
                        }
                        else
                        {
                            MessageBox.Show("串口A未打开！", "Warning");
                        }

                        break;
                    case "电机停机":
                        if (PortA.IsOpen == true)
                        {
                            if (PortB.IsOpen == true)
                            {
                                Byte[] command = new Byte[16];
                                command[0] = 0x66;
                                command[1] = 0x77;                               
                                command[14] = 0x88;
                                command[15] = 0x99;
                                PortA.Write(command, 0, command.Length);
                                PortB.Write(command, 0, command.Length);
                                btnMotorEnable.Text = "电机使能";
                                BtnEnable(false);
                            }
                            else
                            {
                                MessageBox.Show("串口B未打开！", "Warning");
                            }
                        }
                        else
                        {
                            MessageBox.Show("串口A未打开！", "Warning");
                        }

                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
            
        }
        /// <summary>
        /// 按钮使能逻辑
        /// </summary>
        /// <param name="enable"></param>
        public void BtnEnable(bool enable)
        {
            if (enable)
            {
                btn_Listen.Enabled = true;
                btnXForward.Enabled = true;
                btnXBack.Enabled = true;
                btnYForward.Enabled = true;
                btnYBack.Enabled = true;
                btnZForward.Enabled = true;
                btnZBack.Enabled = true;
                btnFForw.Enabled = true;
                btnFBack.Enabled = true;
                btn_Origin.Enabled = true;
            }
            else
            {
                btn_Listen.Enabled = false;
                btnXForward.Enabled = false;
                btnXBack.Enabled = false;
                btnYForward.Enabled = false;
                btnYBack.Enabled = false;
                btnZForward.Enabled = false;
                btnZBack.Enabled = false;
                btnFForw.Enabled = false;
                btnFBack.Enabled = false;
                btn_Origin.Enabled = false;
            }
        }
        /// <summary>
        /// 删去500个点之外多余的点
        /// </summary>
        private void RemoveData()
        {
            X_Axis_PositionList.RemoveAt(0);
            X_Axis_TimeList.RemoveAt(0);
            Y_Axis_PositionList.RemoveAt(0);
            Y_Axis_TimeList.RemoveAt(0);
            Z_Axis_PositionList.RemoveAt(0);
            Z_Axis_TimeList.RemoveAt(0);
            F_Axis_PositionList.RemoveAt(0);
            F_Axis_TimeList.RemoveAt(0);
            TensionForceList.RemoveAt(0);
        }
        /// <summary>
        /// 归零原点操作
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Origin_Click(object sender, EventArgs e)
        {
            if (PortA.IsOpen == true && PortB.IsOpen == true)
            {
                Byte[] command = new Byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x0A;
                command[3] = 0x34;
                command[4] = 0x56;
                command[5] = 0x78;
                command[6] = 0x21;
                command[7] = 0x33;
                command[8] = 0x32;
                command[9] = 0x12;
                command[10] = 0x66;
                command[11] = 0x35;
                command[12] = 0x77;
                command[13] = 0x93;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                Byte[] command1 = new Byte[16];
                command1[0] = 0x66;
                command1[1] = 0x77;
                command1[2] = 0x0B;
                command1[3] = 0x34;
                command1[4] = 0x56;
                command1[5] = 0x78;
                command1[6] = 0x21;
                command1[7] = 0x33;
                command1[8] = 0x32;
                command1[9] = 0x12;
                command1[10] = 0x66;
                command1[11] = 0x35;
                command1[12] = 0x77;
                command1[13] = 0x94;
                command1[14] = 0x88;
                command1[15] = 0x99;
                PortB.Write(command1, 0, command.Length);
            }
            else
            {
                MessageBox.Show("串口未打开！", "Warning");
            }
        }
        /// <summary>
        /// X坐标显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chart_X_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                HitTestResult MyHitTestResult = chart_X.HitTest(e.X, e.Y);
                if (MyHitTestResult.ChartElementType == ChartElementType.DataPoint)
                {
                    this.Cursor = Cursors.Cross;
                    DataPoint dp_X = MyHitTestResult.Series.Points[MyHitTestResult.PointIndex];
                    if (toolStripBtn_Import.Enabled==true)
                    {
                        toolTip_X.SetToolTip(chart_X, "坐标值：" + dp_X.YValues[0] + Environment.NewLine + "时间：" + dp_X.AxisLabel);
                    }
                    else
                    {
                        toolTip_X.SetToolTip(chart_X, "坐标值：" + dp_X.YValues[0] + Environment.NewLine + "时间：" + DateTime.FromOADate(dp_X.XValue).ToString("HH:mm:ss"));    
                    }
                }
                else
                {
                    this.Cursor = Cursors.Default;
                    toolTip_X.Hide(chart_X);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Y坐标显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chart_Y_MouseMove_1(object sender, MouseEventArgs e)
        {
            try
            {
                HitTestResult MyHitTestResult = chart_Y.HitTest(e.X, e.Y);
                if (MyHitTestResult.ChartElementType == ChartElementType.DataPoint)
                {
                    this.Cursor = Cursors.Cross;
                    DataPoint dp_Y = MyHitTestResult.Series.Points[MyHitTestResult.PointIndex];
                    if (toolStripBtn_Import.Enabled == true)
                    {
                        toolTip_Y.SetToolTip(chart_Y, "坐标值：" + dp_Y.YValues[0] + Environment.NewLine + "时间：" + dp_Y.AxisLabel);
                    }
                    else
                    {
                        toolTip_Y.SetToolTip(chart_Y, "坐标值：" + dp_Y.YValues[0] + Environment.NewLine + "时间：" + DateTime.FromOADate(dp_Y.XValue).ToString("HH:mm:ss"));
                    }
                }
                else
                {
                    this.Cursor = Cursors.Default;
                    toolTip_Y.Hide(chart_Y);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// Z坐标显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chart_Z_MouseMove_1(object sender, MouseEventArgs e)
        {
            try
            {
                HitTestResult MyHitTestResult = chart_Z.HitTest(e.X, e.Y);
                if (MyHitTestResult.ChartElementType == ChartElementType.DataPoint)
                {
                    this.Cursor = Cursors.Cross;
                    DataPoint dp_Z = MyHitTestResult.Series.Points[MyHitTestResult.PointIndex];
                    if (toolStripBtn_Import.Enabled == true)
                    {
                        toolTip_Z.SetToolTip(chart_Z, "坐标值：" + dp_Z.YValues[0] + Environment.NewLine + "时间：" + dp_Z.AxisLabel);
                    }
                    else
                    {
                        toolTip_Z.SetToolTip(chart_Z, "坐标值：" + dp_Z.YValues[0] + Environment.NewLine + "时间：" + DateTime.FromOADate(dp_Z.XValue).ToString("HH:mm:ss"));
                    }                    
                }
                else
                {
                    this.Cursor = Cursors.Default;
                    toolTip_Z.Hide(chart_Z);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// F坐标显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chart_F_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                HitTestResult MyHitTestResult = chart_F.HitTest(e.X, e.Y);
                if (MyHitTestResult.ChartElementType == ChartElementType.DataPoint)
                {
                    this.Cursor = Cursors.Cross;
                    DataPoint dp_F = MyHitTestResult.Series.Points[MyHitTestResult.PointIndex];
                    if (toolStripBtn_Import.Enabled == true)
                    {
                        toolTip_F.SetToolTip(chart_F, "坐标值：" + dp_F.YValues[0] + Environment.NewLine + "时间：" + dp_F.AxisLabel);
                    }
                    else
                    {
                        toolTip_F.SetToolTip(chart_F, "坐标值：" + dp_F.YValues[0] + Environment.NewLine + "时间：" + DateTime.FromOADate(dp_F.XValue).ToString("HH:mm:ss"));
                    }
                }
                else
                {
                    this.Cursor = Cursors.Default;
                    toolTip_F.Hide(chart_F);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void 返回上一级缩放ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chart_X.ChartAreas[0].AxisX.ScaleView.ZoomReset();
            chart_X.ChartAreas[0].AxisY.ScaleView.ZoomReset();

            chart_Y.ChartAreas[0].AxisX.ScaleView.ZoomReset();
            chart_Y.ChartAreas[0].AxisY.ScaleView.ZoomReset();

            chart_Z.ChartAreas[0].AxisX.ScaleView.ZoomReset();
            chart_Z.ChartAreas[0].AxisY.ScaleView.ZoomReset();

            chart_F.ChartAreas[0].AxisX.ScaleView.ZoomReset();
            chart_F.ChartAreas[0].AxisY.ScaleView.ZoomReset();
        }
        /// <summary>
        /// Listning Start
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Listen_Click(object sender, EventArgs e)
        {
            switch (btn_Listen.Text)
            {
                case "开始监听":
                    timer_Listen.Interval = 500;
                    timer_Listen.Start();
                    btn_Listen.Text = "结束监听";
                    break;
                case "结束监听":
                    timer_Listen.Stop();
                    btn_Listen.Text = "开始监听";
                    break;
                default:
                    break;
            }
            timer_Listen.Interval = 500;
            timer_Listen.Enabled = true;
        }
        /// <summary>
        /// 定时监听
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer_Listen_Tick(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0xCC;
                command[3] = 0xCC;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// X_Axis go forward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnXForward_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x01;
                command[3] = 0x33;
                command[4] = 0x0A;
                command[5] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                X_Axis_RelativeDisplacement+=10;
                lb_XAxis.Text = "X轴位移：" + X_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// Y_Axis go forward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnYForward_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x02;
                command[3] = 0x33;
                command[6] = 0x0A;
                command[7] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                Y_Axis_RelativeDisplacement += 10;
                lb_YAxis.Text = "Y轴位移：" + Y_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// Z_Axis go forward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnZForward_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x03;
                command[3] = 0x33;
                command[8] = 0x0A;
                command[9] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                Z_Axis_RelativeDisplacement += 10;
                lb_ZAxis.Text = "Z轴位移：" + Z_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// F_Axis go forward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFForw_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x04;
                command[3] = 0x33;
                command[10] = 0x0A;
                command[11] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                F_Axis_RelativeDisplacement += 10;
                lb_FAxis.Text = "F轴位移：" + F_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// X_Axis go backward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnXBack_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x01;
                command[3] = 0x44;
                command[4] = 0x0A;
                command[5] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                X_Axis_RelativeDisplacement -= 10;
                lb_XAxis.Text = "X轴位移：" + X_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// Y_Axis go backward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnYBack_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x02;
                command[3] = 0x44;
                command[6] = 0x0A;
                command[7] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                Y_Axis_RelativeDisplacement -= 10;
                lb_YAxis.Text = "Y轴位移：" + Y_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// Z_Axis go backward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnZBack_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x03;
                command[3] = 0x44;
                command[8] = 0x0A;
                command[9] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                Z_Axis_RelativeDisplacement -= 10;
                lb_ZAxis.Text = "Z轴位移：" + Z_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// F_Axis go backward
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFBack_Click(object sender, EventArgs e)
        {
            try
            {
                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x04;
                command[3] = 0x44;
                command[10] = 0x0A;
                command[11] = 0x0A;
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
                F_Axis_RelativeDisplacement -= 10;
                lb_FAxis.Text = "F轴位移：" + F_Axis_RelativeDisplacement.ToString() + "mm";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Warning");
            }
        }
        /// <summary>
        /// Mannnual Operation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAutoSend_Click(object sender, EventArgs e)
        {
            try
            {
                int XCoordinate = Convert.ToInt32(tbx_XCoordinate.Text);
                int YCoordinate = Convert.ToInt32(tbx_YCoordinate.Text);
                int ZCoordinate = Convert.ToInt32(tbx_ZCoordinate.Text);
                if (XCoordinate>2000)
                {
                    errorProviderX.SetError(tbx_XCoordinate, "请输入-2000~2000内的整数");
                    return;
                }
                else
                {
                    errorProviderX.SetError(tbx_XCoordinate, null);
                }
                if (YCoordinate > 2000)
                {
                    errorProviderY.SetError(tbx_YCoordinate, "请输入-2000~2000内的整数");
                    return;
                }
                else
                {
                    errorProviderY.SetError(tbx_YCoordinate, null);
                }
                if (ZCoordinate > 2000)
                {
                    errorProviderZ.SetError(tbx_ZCoordinate, "请输入-2000~2000内的整数");
                    return;
                }
                else
                {
                    errorProviderZ.SetError(tbx_ZCoordinate, null);
                }
                byte[] xbuffer = BitConverter.GetBytes(XCoordinate);                    //低字节转化后在低位，共四个字节,只用1/0两字节
                byte[] ybuffer = BitConverter.GetBytes(YCoordinate);
                byte[] zbuffer = BitConverter.GetBytes(ZCoordinate);

                byte[] command = new byte[16];
                command[0] = 0x66;
                command[1] = 0x77;
                command[2] = 0x22;
                command[3] = 0x22;
                command[4] = xbuffer[1];
                command[5] = xbuffer[0];
                command[6] = ybuffer[1];
                command[7] = ybuffer[0];
                command[8] = zbuffer[1];
                command[9] = zbuffer[0];
                command[14] = 0x88;
                command[15] = 0x99;
                PortA.Write(command, 0, command.Length);
                PortB.Write(command, 0, command.Length);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(),"Warning");
            }
            
        }
        /// <summary>
        /// 导出Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripBtn_Export_Click(object sender, EventArgs e)
        {
            XSSFWorkbook ExportExcelWorkbook = this.SaveDataToExcel();
            FileStream fs=null;
            string ExcelFileName = "";
            try
            {
                saveFileDialog1.Filter = "Excel文件（*.xlsx）|*.xlsx";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ExcelFileName = saveFileDialog1.FileName.TrimEnd((".xlsx").ToCharArray())+"_"+DateTime.Now.ToString("yyyy-MM-dd")+".xlsx";
                    ISheet sheet1 = ExportExcelWorkbook.GetSheetAt(0);
                    for (int i = 1; i < X_Axis_PositionList.Count; i++)
                    {
                        IRow dataRow= sheet1.CreateRow(i);
                        dataRow.CreateCell(0).SetCellValue(X_Axis_TimeList[i - 1]);
                        dataRow.CreateCell(1).SetCellValue(X_Axis_PositionList[i - 1]);
                        dataRow.CreateCell(2).SetCellValue(Y_Axis_PositionList[i - 1]);
                        dataRow.CreateCell(3).SetCellValue(Z_Axis_PositionList[i - 1]);
                        dataRow.CreateCell(4).SetCellValue(F_Axis_PositionList[i - 1]);
                        dataRow.CreateCell(5).SetCellValue(TensionForceList[i - 1]);
                    }
                    using (fs=new FileStream(ExcelFileName,FileMode.Create))
                    {
                        ExportExcelWorkbook.Write(fs);
                        fs.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                if (fs!=null)
                {
                    fs.Close();
                }
                MessageBox.Show(ex.ToString());
            }          
        }
        /// <summary>
        /// 导入Excel并显示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripBtn_Import_Click(object sender, EventArgs e)
        {
            FileStream fs = null;
            string ExcelPath = null;
            DataRow datarow = null;
            DataColumn datacolumn = new DataColumn();
            DataTable dataTable=new DataTable();
            openFileDialog1.Title = "打开Excel";
            openFileDialog1.Filter = "所有文件（*.*）|*.*|Excel文件（*.xls）|*.xls";
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr==DialogResult.OK)
            {
                ExcelPath = openFileDialog1.FileName;
                try
                {
                    using (fs = new FileStream(ExcelPath, FileMode.Open))
                    {
                        XSSFWorkbook wb = new XSSFWorkbook(fs);
                        for (int i = 0; i < wb.NumberOfSheets; i++)
                        {
                            ISheet ExcelSheet = wb.GetSheetAt(i);
                            IRow FirstRow = ExcelSheet.GetRow(0);
                            for (int c = 0; c < FirstRow.LastCellNum; c++)
                            {
                                ICell FirstRowCell = FirstRow.GetCell(c);
                                datacolumn = new DataColumn(FirstRowCell.ToString());
                                dataTable.Columns.Add(datacolumn);
                            }
                            for (int j = 1; j < ExcelSheet.LastRowNum+1; j++)             //第一行是列名，从第二行开始
                            {
                                IRow row = ExcelSheet.GetRow(j);
                                datarow = dataTable.NewRow();
                                for (int k = 0; k < row.LastCellNum; k++)               
                                {
                                    ICell cell = row.GetCell(k);
                                    if (HSSFDateUtil.IsCellDateFormatted(cell))
                                    {
                                        datarow[k] = cell.DateCellValue.ToString("HH:mm:ss");
                                        continue;
                                    }
                                    datarow[k] = cell.NumericCellValue;
                                }
                                dataTable.Rows.Add(datarow);
                            }                          
                        }
                        fs.Close();
                    }
                    chart_X.DataSource=dataTable;
                    chart_X.Series[0].XValueMember="时间";
                    chart_X.Series[0].YValueMembers = "X轴坐标";
                    chart_X.Series[0].IsValueShownAsLabel=true;
                    chart_X.DataBind();

                    chart_Y.DataSource = dataTable;
                    chart_Y.Series[0].XValueMember = "时间";
                    chart_Y.Series[0].YValueMembers = "Y轴坐标";
                    chart_Y.Series[0].IsValueShownAsLabel = true;
                    chart_Y.DataBind();

                    chart_Z.DataSource = dataTable;
                    chart_Z.Series[0].XValueMember = "时间";
                    chart_Z.Series[0].YValueMembers = "Z轴坐标";
                    chart_Z.Series[0].IsValueShownAsLabel = true;
                    chart_Z.DataBind();

                    chart_F.DataSource = dataTable;
                    chart_F.Series[0].XValueMember = "时间";
                    chart_F.Series[0].YValueMembers = "F轴坐标";
                    chart_F.Series[0].IsValueShownAsLabel = true;
                    chart_F.DataBind();
                }
                catch (Exception ex)
                {
                    if (fs!=null)
                    {
                        fs.Close();
                    }
                    MessageBox.Show(ex.ToString(), "Warning");
                }                
            }
        }
        /// <summary>
        /// 第一行作为各列的表头
        /// </summary>
        /// <returns></returns>
        private XSSFWorkbook SaveDataToExcel()
        {
            XSSFWorkbook ExportDataWorkbook = new XSSFWorkbook();
            ISheet ExportDataSheet = ExportDataWorkbook.CreateSheet("Sheet1");
            IRow ExportDataFirstRow = ExportDataSheet.CreateRow(0);
            ExportDataFirstRow.CreateCell(0).SetCellValue("时间");
            ExportDataFirstRow.CreateCell(1).SetCellValue("X轴坐标");
            ExportDataFirstRow.CreateCell(2).SetCellValue("Y轴坐标");
            ExportDataFirstRow.CreateCell(3).SetCellValue("Z轴坐标");
            ExportDataFirstRow.CreateCell(4).SetCellValue("F轴坐标");
            ExportDataFirstRow.CreateCell(5).SetCellValue("F轴坐标");
            ExportDataFirstRow.CreateCell(6).SetCellValue("力传感器数据");
            return ExportDataWorkbook;
        }
        /// <summary>
        /// 关于界面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 关于AToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form fm = new Form2();
            fm.Owner = this;
            fm.StartPosition = FormStartPosition.CenterScreen;
            fm.ShowDialog(this);
        }
        /// <summary>
        /// 帮助文档界面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 帮助文档HToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form fm = new Form3();
            fm.Owner = this;
            fm.StartPosition = FormStartPosition.CenterScreen;
            fm.ShowDialog(this);
        }
        /// <summary>
        /// 记录最近五个点的信息，绘制3D图用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRecord_X_Click(object sender, EventArgs e)
        {
            
        }

        private void btnRecord_X_Click_1(object sender, EventArgs e)
        {
        
        }
      
    }
}
