
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;

namespace MqttToTTS
{
    public partial class Form1 : Form
    {
        private MqttClient mqttClient = null;
        private string clientId = null;

        string speaker = "";
        int ttsRate = 0;
        int ttsVol = 80;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                textBox1.Text = AppConfig.GetValue("Host", "");
                textBox2.Text = AppConfig.GetValue("Topic","");
                int ttsIsOpen = int.Parse(AppConfig.GetValue("TTS","0"));
                int shutDownIsOpen = int.Parse(AppConfig.GetValue("ShutDown","0"));
                //comboBox1.Items.AddRange((object[])synth.GetInstalledVoices());
                //初始化TTS
                SpeechSynthesizer synth = new SpeechSynthesizer();
                var speakers = synth.GetInstalledVoices();
                comboBox1.Text = speakers[0].VoiceInfo.Name;
                foreach (var voice in speakers)
                {
                    comboBox1.Items.Add(voice.VoiceInfo.Name);
                }
                speaker = speakers[0].VoiceInfo.Name;
                //初始化MQTT连接
                string topic = textBox2.Text;
                string host = textBox1.Text;
                clientId = Guid.NewGuid().ToString();
                if (!string.IsNullOrEmpty(host) && !string.IsNullOrEmpty(topic))
                {
                    // 实例化Mqtt客户端 
                    mqttClient = new MqttClient(host);
                    // 注册接收消息事件 
                    mqttClient.MqttMsgPublishReceived += client_MqttMsgPublishReceived;

                    
                    mqttClient.Connect(clientId);

                    button1.Text = "断开";
                    label4.Text = clientId;
                    // 订阅主题 "/home/temperature"， 订阅质量 QoS 2 
                    mqttClient.Subscribe(new string[] { topic }, new byte[] { MqttMsgBase.QOS_LEVEL_AT_MOST_ONCE });
                }
                if (ttsIsOpen != 0) {
                    checkBox1.Checked = true;
                }
                if (shutDownIsOpen != 0) {
                    checkBox2.Checked = true;
                }

                if (DesignMode)
                {
                    notifyIcon1.Visible = false;
                    return;
                }
                this.notifyIcon1.Text = this.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        void client_MqttMsgPublishReceived(object sender, MqttMsgPublishEventArgs e)
        {
            if (checkBox1.Checked)
            {
                string text = Encoding.UTF8.GetString(e.Message);
                if (text == "ShutDown")
                {
                    text = "此设备即将关机";
                }
                using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
                {
                    synthesizer.SelectVoice(speaker);
                    synthesizer.Rate = ttsRate;
                    synthesizer.Volume = ttsVol;
                    synthesizer.Speak(text);
                    synthesizer.Dispose();
                    GC.Collect();
                }
            }
            if (checkBox2.Checked)
            {
                if (Encoding.UTF8.GetString(e.Message) == "ShutDown")
                {
                    Shutdown();
                }
            }
            // 打印订阅的发布端消息
            //Console.WriteLine(string.Format("subscriber,topic:{0},content:{1}", e.Topic, Encoding.UTF8.GetString(e.Message)));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (button1.Text == ("连接"))
                {
                    AppConfig.UpdateAppConfig("Host", textBox1.Text);
                    //ConfigurationManager.AppSettings["Host"] = textBox1.Text;
                    AppConfig.UpdateAppConfig("Topic", textBox2.Text);
                    //ConfigurationManager.AppSettings["Topic"] = textBox2.Text;
                    if (checkBox1.Checked)
                    {
                        AppConfig.UpdateAppConfig("TTS" , "1");
                    }
                    else {
                        AppConfig.UpdateAppConfig("TTS","0");
                    }
                    if (checkBox2.Checked)
                    {
                        AppConfig.UpdateAppConfig("ShutDown", "1");
                    }
                    else {
                        AppConfig.UpdateAppConfig("ShutDown", "0");
                    }
                    ttsVol = trackBar1.Value;
                    speaker = comboBox1.Text;
                    ttsRate = int.Parse(domainUpDown1.Text);
                    string topic = textBox2.Text;
                    string host = textBox1.Text;
                    // 实例化Mqtt客户端 
                    mqttClient = new MqttClient(host);
                    mqttClient.MqttMsgPublishReceived += client_MqttMsgPublishReceived;
                    mqttClient.Subscribe(new string[] { topic }, new byte[] { MqttMsgBase.QOS_LEVEL_AT_MOST_ONCE });
                    mqttClient.Connect(clientId);
                    button1.Text = "断开";
                }
                else
                {
                    mqttClient.Disconnect();
                    button1.Text = "连接";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        #region 界面显示相关
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            From_Visible(true);
        }
        private void From_Visible(Boolean visible)
        {
            if (visible == true)
            {
                this.Visible = true;
                this.WindowState = FormWindowState.Normal;
                int x = Screen.PrimaryScreen.WorkingArea.Width / 2 - this.Width / 2;
                int y = Screen.PrimaryScreen.WorkingArea.Height / 2 - this.Height / 2;
                this.Location = new Point(x, y);
                this.Focus();
            }
            else
            {
                this.Visible = false;
            }
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                From_Visible(false);
            }
            else
            {
                From_Visible(true);
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (DesignMode == false)
            {
                From_Visible(false);
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("是否关闭程序，请确认", "确认退出", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.OK)
            {
                notifyIcon1.Visible = false;
                Environment.Exit(0);
            }
            else
            {
                e.Cancel = true;
            }
        }
        #endregion

        //执行命令
        public void Exec(string str)
        {
            try
            {
                using (Process process = new Process())
                {
                    process.StartInfo.FileName = "cmd.exe";//调用cmd.exe程序
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardInput = true;//重定向标准输入
                    process.StartInfo.RedirectStandardOutput = true;//重定向标准输出
                    process.StartInfo.RedirectStandardError = true;//重定向标准出错
                    process.StartInfo.CreateNoWindow = true;//不显示黑窗口
                    process.Start();//开始调用执行
                    process.StandardInput.WriteLine(str + "&exit");//标准输入str + "&exit"，相等于在cmd黑窗口输入str + "&exit"
                    process.StandardInput.AutoFlush = true;//刷新缓冲流，执行缓冲区的命令，相当于输入命令之后回车执行
                    process.WaitForExit();//等待退出
                    process.Close();//关闭进程
                }
            }
            catch
            {
            }
        }
        //执行关机操作
        public void Shutdown()
        {
            this.Exec("shutdown -s -f -t 0");
        }
    }
}
