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
using System.Threading;

namespace yankong
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Updata_Serialport_Name(ComboBox MycomboBox)
        {
            string[] ArryPort;                               // 定义字符串数组，数组名为 ArryPort
            ArryPort = SerialPort.GetPortNames();            // SerialPort.GetPortNames()函数功能为获取计算机所有可用串口，以字符串数组形式输出
            MycomboBox.Items.Clear();                        // 清除当前组合框下拉菜单内容                  
            for (int i = 0; i < ArryPort.Length; i++)
            {
                MycomboBox.Items.Add(ArryPort[i]);           // 将所有的可用串口号添加到端口对应的组合框中
            }
        }
        private bool SetSpeed(int speed, int waitTime = 200)
        {
            try
            {
                var speedBytes = BitConverter.GetBytes(speed);
                var bytes = new byte[8];
                bytes[0] = 0x01;
                bytes[1] = 0x06; //功能码
                bytes[2] = 0x00; //寄存器地址
                bytes[3] = 0x23; //寄存器地址
                bytes[4] = speedBytes[1]; //写入数据
                bytes[5] = speedBytes[0]; //写入数据

                var tmpData = new List<byte>();
                for (int i = 0; i <= 5; i++)
                {
                    tmpData.Add(bytes[i]);
                }
                byte dl, dh;
                CalCRC(tmpData.ToArray(), out dl, out dh);
                bytes[6] = dl;//1-6字节累加校正码，低
                bytes[7] = dh;//1-6字节累加校正码，高

                for (var i = 0; i < 200; i++)
                {
                    serialPort1.Write(bytes, 0, bytes.Length);
                    Thread.Sleep(waitTime);
                    for (var j = 0; j < 200; j++)
                    {
                        var backData = new byte[serialPort1.BytesToRead]; //读取
                        if (serialPort1.Read(backData, 0, backData.Length) < 1) continue;
                        if (backData.Length < 8) continue;

                        bool flag = true;
                        for (int k = 0; k < 8; k++)
                        {
                            if (backData[k] != bytes[k])
                            {
                                flag = false;
                            }
                        }
                        if (!flag) continue;

                        //发出和返回完全一致，则通过
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        public bool Start(int speed, bool fwd)
        {
            speed = fwd ? speed : -speed; //速度正数正向，负数负向
            if (!SetSpeed(speed)) return false;

            try
            {
                var bytes = new byte[8];
                bytes[0] = 0x01; //地址
                bytes[1] = 0x06; //功能码
                bytes[2] = 0x00; //寄存器地址
                bytes[3] = 0x27; //寄存器地址
                bytes[4] = 0x00; //数据
                bytes[5] = 0x02; //数据  //速度模式

                var tmpData = new List<byte>();
                for (int i = 0; i <= 5; i++)
                {
                    tmpData.Add(bytes[i]);
                }
                byte dl, dh;
                CalCRC(tmpData.ToArray(), out dl, out dh);
                bytes[6] = dl;//1-6字节累加校正码，低
                bytes[7] = dh;//1-6字节累加校正码，高

                
                for (var i = 0; i < 200; i++)
                {
                    serialPort1.Write(bytes, 0, bytes.Length);
                    Thread.Sleep(200);
                    for (var j = 0; j < 200; j++)
                    {
                        var backData = new byte[serialPort1.BytesToRead]; //读取
                        if (serialPort1.Read(backData, 0, backData.Length) < 1) continue;
                        if (backData.Length < 8) continue;

                        bool flag = true;
                        for (int k = 0; k < 8; k++)
                        {
                            if (backData[k] != bytes[k])
                            {
                                flag = false;
                            }
                        }
                        if (!flag) continue;
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        //启动电机
        private void button1_Click(object sender, EventArgs e)
        {
            bool fx ;
            if (radioButton1.Checked)
            {
                fx = true;
            }
            else
            {
                fx = false;
            }
            Start(int.Parse(textBox1.Text),fx);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Updata_Serialport_Name(comboBox1); // 调用更新可用串口函数，comboBox1为端口号组合框的名称
        }

        //打开串口
        private void button3_Click(object sender, EventArgs e)
        {
            try                                                          
            {
                serialPort1.PortName = comboBox1.Text;                   // 将串口设备的串口号属性设置为comboBox1复选框中选择的串口号
                serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);  // 将串口设备的波特率属性设置为comboBox2复选框中选择的波特率
                serialPort1.Open();                                      // 打开串口
                comboBox1.Enabled = false;                               // 串口已打开，将comboBox1、comboBox2设置为不可操作
                comboBox2.Enabled = false;                               
                button3.Enabled = false;                                 //打开串口后，此按钮不可操作
                button4.Enabled = true;
            }
            catch
            {
                MessageBox.Show("打开串口失败，请检查串口", "错误");     // 弹出错误对话框
            }
        }

        //关闭串口
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                serialPort1.Close();                     // 关闭串口
                comboBox1.Enabled = true;                // 串口已关闭，将comboBox1设置为可操作
                comboBox2.Enabled = true;                // 串口已关闭，将comboBox2设置为可操作
                button3.Enabled = true;
                button4.Enabled = false;
            }
            catch
            {
                MessageBox.Show("关闭串口失败，请检查串口", "错误");   // 弹出错误对话框
            }
        }

        //关闭电机
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                var bytes = new byte[8];
                bytes[0] = 0x01;
                bytes[1] = 0x06; //功能码
                bytes[2] = 0x00; //寄存器地址
                bytes[3] = 0x28; //寄存器地址
                bytes[4] = 0x00; //数据
                bytes[5] = 0x00; //数据

                var tmpData = new List<byte>();
                for (int i = 0; i <= 5; i++)
                {
                    tmpData.Add(bytes[i]);
                }
                byte dl, dh;
                CalCRC(tmpData.ToArray(), out dl, out dh);
                bytes[6] = dl;//1-6字节累加校正码，低
                bytes[7] = dh;//1-6字节累加校正码，高

                serialPort1.Write(bytes, 0, bytes.Length);
            }
            catch (Exception ex)
            {
                MessageBox.Show("关闭电机失败！", "错误");   // 弹出错误对话框
            }
        }

        //CRC校验函数
        public static void CalCRC(byte[] bytes, out byte dl, out byte dh)
        {
            int crc = 0xffff;
            int len = bytes.Length;
            for (int n = 0; n < len; n++)
            {
                byte i;
                crc = crc ^ bytes[n];
                for (i = 0; i < 8; i++)
                {
                    int TT;
                    TT = crc & 1;
                    crc = crc >> 1;
                    crc = crc & 0x7fff;
                    if (TT == 1)
                    {
                        crc = crc ^ 0xa001;
                    }
                    crc = crc & 0xffff;
                }

            }
            dl = (byte)((crc & 0xff));
            dh = (byte)((crc >> 8) & 0xff);
        }
    }
}
