using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports; //COMポートを使うためです。
using System.Text;     
using Ivi.Visa.Interop; //USBを使うためです。
using Aspose.Cells;
using System.Runtime.ConstrainedExecution;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.InteropServices;
using System.IO;


namespace MO_test9
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            Function function = new Function("COM3", "USB::0x0B3E::0x104A::CP002893::INSTR", "GPIB::6::INSTR");


            //function.Workbook_set();


            //function.Device_open();
            //function.Device_close();

            //Data data1 = function.Measurement_main();

            //double aaaa = data1.mag;
            //Console.WriteLine(data1.mag);
            //Console.WriteLine(data1.faraday_deg);
        }
    }

    public class Data
    {
        public double mag;
        public double faraday_deg;
    }


    class Function
    {
        string comNum; //COMポートの番号です。
        string pInst_magpow; //電磁石電源のインスタンスパスです。
        string pInst_multi; //マルチメータのインスタンスパスです。

        SerialPort serialPort;
        ResourceManager rm_magpow = new ResourceManager();
        FormattedIO488 msg_magpow = new FormattedIO488();
        ResourceManager rm_multi = new ResourceManager();
        FormattedIO488 msg_multi = new FormattedIO488();

        Workbook workbook = new Workbook();

        double current_theta;
        double nonmedia_theta;

        public Function(string comNum, string pInst_magpow, string pInst_multi)
        {   
            //コンストラクタを書くところです。
            this.comNum = comNum;
            this.pInst_magpow = pInst_magpow;
            this.pInst_multi = pInst_multi;

            //  COMポートの設定をします。
            string portName = comNum; // 使用するCOMポート名を指定してください"COM3"
            int baudRate = 9600;
            Parity parity = Parity.None;
            int dataBits = 8;
            StopBits stopBits = StopBits.One;
            // シリアルポートのインスタンスを作成
            serialPort = new SerialPort(portName, baudRate, parity, dataBits, stopBits);
            serialPort.ReadTimeout = 1500;
            serialPort.WriteTimeout = 1500;
            serialPort.Handshake = Handshake.None;  //ハンドシェイク
            serialPort.Encoding = Encoding.UTF8;          //エンコード
            serialPort.NewLine = "\r";                   //改行コード指定
        }


        //セットアップ用の関数です。
        public void Device_open()
        {
            //  COMポートを開きます。
            serialPort.Open();

            //IDN?ここは消しても問題ないです。
            serialPort.WriteLine("*IDN?");

            string response = serialPort.ReadLine();

            Console.WriteLine(response);
            //

            //インスタンスパスで開く機器を開きます。
            //電磁石電源です。"USB::0x0B3E::0x104A::CP002893::INSTR"
            msg_magpow.IO = (IMessage)rm_magpow.Open(pInst_magpow, AccessMode.NO_LOCK, 0, "");

            //IDN?
            msg_magpow.WriteString("*IDN?");

            response = msg_magpow.ReadString();

            Console.WriteLine(response);
            //

            //マルチメータです。"GPIB::6::INSTR"
            msg_multi.IO = (IMessage)rm_multi.Open(pInst_multi, AccessMode.NO_LOCK, 0, "");

            //IDN?
            msg_multi.WriteString("*IDN?");

            response = msg_multi.ReadString();

            Console.WriteLine(response);
            //
        }

        public void Device_close()
        {
            serialPort.Close();
            msg_magpow.IO.Close();
            msg_multi.IO.Close();
        }

        public void Workbook_set()
        {
            Worksheet worksheet1 = workbook.Worksheets[0];
            Worksheet worksheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            worksheet1.Name = "Data";
            worksheet2.Name = "Graph";

            worksheet1.Cells[0, 0].PutValue("H[mT]");
            worksheet1.Cells[0, 1].PutValue("Faraday deg[deg]");


            workbook.Save("MO_data.xlsx");
        }

        public void Theta_read()
        {
            //textからThetaを読み込みます

            string path = @"C:\Users\yohei\Downloads\MO_test9\MO_test9\datalist.txt";

            StreamReader sr = new StreamReader(path, Encoding.GetEncoding("Shift_JIS"));

            current_theta = double.Parse(sr.ReadLine());   //doubleにstringから変換します
            nonmedia_theta = double.Parse(sr.ReadLine());
            
            sr.Close();
        }

        public void Theta_write()
        {
            string path = @"C:\Users\yohei\Downloads\MO_test9\MO_test9\datalist.txt";

            StreamWriter sw = new StreamWriter(path, false);

            sw.WriteLine(current_theta);
            sw.WriteLine(nonmedia_theta);

            sw.Close();
        }

        //磁石を動かすやつです。
        public double Mag_output(double h_target, Boolean flag_reverse)
        {
            double cur = 0.00378515 * h_target - 0.02470835;
            cur = Math.Round(cur, 3, MidpointRounding.AwayFromZero);

            Convert.ToString(cur);

            msg_magpow.WriteString("CURR " + cur);
            msg_magpow.WriteString("OUTP 1");

            if (flag_reverse == false)
            {
                return (264.19 * cur + 6.5277);
            }
            else
            {
                return (-264.19 * cur + 6.5277);
            }

        }


        //ここから測定関数です。
        public Data Measurement_main(double start_theta) //引数は、測定の開始角度を入れます
        {// 最下点のファラデー回転角を磁界とともに返すメイン測定関数です。

            var datalist_intensity = new List<double>();
            var datalist_theta = new List<double>();
            double dtheta = 250;
            int data_collectflag = 0;

            Boolean start_flag = false;
            while (start_flag == false){
                //スタート地点から上下をはかり、どちらが減少傾向にあるか確かめます。
                //もしも同じ値が出たら、dthetaを小さくしてもう一回します。
                current_theta = start_theta;

                datalist_intensity.Add(Measurement_read(0));
                datalist_theta.Add(current_theta);

                datalist_intensity.Add(Measurement_read(dtheta));
                datalist_theta.Add(current_theta);

                current_theta = start_theta;

                datalist_intensity.Add(Measurement_read(-dtheta));
                datalist_theta.Add(current_theta);

                if(datalist_intensity[1] - datalist_intensity[2] != 0)
                {
                    if (datalist_intensity[1] - datalist_intensity[2] > 0)
                    {
                        dtheta = -dtheta;
                    }
                    start_flag = true;
                }
                else
                {
                    dtheta /= 2;
                }
            }

            //決めた方向に向けて測定をします。
            Boolean revflag = false;
            int datanum = 0;
            int pre_datanum = 0;
            while (data_collectflag< 5 || datalist_intensity.Count < 20)
            {
                datalist_intensity.Add(Measurement_read(dtheta));
                datalist_theta.Add(current_theta);

                datanum = datalist_intensity.Count - 1;
                pre_datanum = datanum - 1;

                if (data_collectflag < 5 && datalist_intensity[datanum] - datalist_intensity[pre_datanum] > 0)
                {
                    data_collectflag ++;
                }

                if (data_collectflag >= 5 && revflag == false)
                {
                    current_theta += dtheta/2;
                    dtheta = -dtheta;
                    revflag = true;
                }
            }

            //測定した中で、最小の点を探します。
            double data_min = datalist_intensity[0];
            datanum = 0;
            int i = 0;
            foreach(double data_compare in datalist_intensity)
            {
                if(data_min >= data_compare)
                {
                    data_min = data_compare;
                    datanum = i;
                }
                i++;
            }

            //最小の点から左右に20点ずつ細かくとります。
            current_theta = datalist_theta[datanum];
            dtheta = 25;

            for(i=0; i<=39; i++)
            {
                datalist_intensity.Add(Measurement_read(dtheta));
                datalist_theta.Add(current_theta);

                if(i == 19)
                {
                    current_theta = datalist_theta[datanum];
                    dtheta = -dtheta;
                }
            }
            
            //データを取り終わりました。近似曲線を出して、最下点を推定します。





            Data data = new Data();
            data.mag = 1;
            data.faraday_deg = 5;


            return data;
        }

        public double Measurement_read(double dtheta)//動かしたい角度を要求して、回して、その先のマルチメータの電圧を返します。
        {
            //角度を足します。
            current_theta += dtheta;
            serialPort.WriteLine("AXI1:GOABS " + current_theta.ToString());
            Delay();

            //マルチメータに値を尋ねます。
            msg_multi.WriteString("*IDN?");
            string response = msg_multi.ReadString();
            string[] response_arry = response.Split('V');

            return (double.Parse(response_arry[0]));
        }

        public void Delay()
        {
            double getpos=0, pre_getpos=1;

            while (getpos - pre_getpos != 0){

                pre_getpos = getpos;
                serialPort.WriteLine("AXI1:POS?");
                string response = serialPort.ReadLine();
                getpos = double.Parse(response);

                if(getpos - pre_getpos == 1)
                {
                    pre_getpos = 0;
                }
            }
        }

    }
}

