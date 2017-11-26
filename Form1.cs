using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Threading;
using System.Diagnostics;
using System.Net;
using System.Globalization;
//using Microsoft.Speech;
//using Microsoft.Speech.Synthesis;
//using Microsoft.Speech.Recognition;



namespace SpeechRecognitionApp {
    
    public partial class Form1 : Form {

        string currentDir = System.AppDomain.CurrentDomain.BaseDirectory;
        //Form Declaration
        SpeechSynthesizer ss = new SpeechSynthesizer();
        PromptBuilder pb = new PromptBuilder();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();
        Choices clist = new Choices(new string[] { "hello", "how are you", "what is the current time", "open chrome",
                                    "thank you", "close",
                                    "get current date",
                                    "get battery status",
                                    "get IP address", "check Internet connection",
                                    "play youtube video number one", "play youtube video number two", "play youtube video number three",
                                    "open windows media player", "open paint",
                                    "get bitcoin value",
                                    "tell currency-exchange", "tell currency-exchange to Dollar"});

        public Form1() {
            
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            //start
           
                btnStart.Enabled = false;
                btnStop.Enabled = true;

            
                //clist.Add(new string[] { "hello", "how are you", "what is the current time", "open chrome",
                //"thank you", "close", "get current date",
                //                        "get battery status", "get IP address", "check Internet connection", 
                //                        "open windows media player", "open paint", "get bitcoin value",
                //                        "tell currency-exchange", "tell currency-exchange to Dollar"});

                Grammar gr = new Grammar(new GrammarBuilder(clist));

                try {
                    sre.RequestRecognizerUpdate();
                    sre.LoadGrammar(gr);
                    sre.SpeechRecognized += sre_SpeechRecognized;
                    sre.SetInputToDefaultAudioDevice();
                    sre.RecognizeAsync(RecognizeMode.Multiple);
                } catch (Exception ex) {
                    MessageBox.Show(ex.Message, "error");
                }
            
        }

        private void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e) {
            
            switch (e.Result.Text.ToString()) {
                case "hello":
                    ss.SpeakAsync("Hello");
                    break;
                case "what is the current time":
                    ss.SpeakAsync("current time is " + DateTime.Now.ToLongTimeString());
                    txtContents.Text += DateTime.Now.ToLongTimeString() + Environment.NewLine; 
                    break;
                case "thank you":
                    ss.SpeakAsync("my pleasure");
                    break;
                case "open chrome":
                    Process.Start("chrome", "https://www.youtube.com/watch?v=xug793JfqSE");

                    break;
                case "close":
                    Application.Exit();
                    break;
                case "get current date":
                    ss.SpeakAsync(getCurrentDate());
                    txtContents.Text += getCurrentDate() + Environment.NewLine;
                    break;
                case "get IP address":
                    ss.SpeakAsync(GetLocalIPAddress());
                    txtContents.Text += GetLocalIPAddress() + Environment.NewLine;
                    break;
                case "check Internet connection":
                    ss.SpeakAsync(CheckIfConnected() + "");
                    txtContents.Text += CheckIfConnected() + Environment.NewLine;
                    break;
                case "open windows media player":
                    Process.Start("wmplayer.exe");
                    //Process.Start("wmplayer.exe", @"C:\Test.mp3");
                    break;
                case "open paint":
                    Process.Start("mspaint.exe");
                    //Process.Start("mspaint.exe", @"C:\Test.jpg");
                    break;
                case "play youtube video number one":
                    Process.Start("chrome", "https://www.youtube.com/watch?v=l-dYNttdgl0");
                    break;
                case "play youtube video number two":
                    Process.Start("chrome", "https://www.youtube.com/watch?v=VC3qO2V1AXY");
                    break;
                case "play youtube video number three":
                    Process.Start("chrome", "https://www.youtube.com/watch?v=H7hGiZ579cs");
                    break;
                case "get bitcoin value":
                    txtContents.Text += "Bitcoin + ";
                    string bitcoinValue = ReturnCurrency("https://coinmarketcap.com/currencies/bitcoin/", currentDir + @"\bitcoinFile.html", "quote_price", 2);
                    ss.SpeakAsync(bitcoinValue);
                    txtContents.Text += bitcoinValue + Environment.NewLine;
                    break;
                case "tell currency-exchange":
                    txtContents.Text += "1 Eur To Ron : ";
                    string valueEurtoRon = ReturnCurrency("http://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=RON", currentDir + @"\EurToRon.html", "uccResultAmount", 120);
                    ss.SpeakAsync(valueEurtoRon);
                    txtContents.Text += valueEurtoRon + Environment.NewLine;
                    break;
                case "tell currency-exchange to Dollar":
                    txtContents.Text += "1 Eur To USD : ";
                    string valueEurToUsd = ReturnCurrency("http://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=USD", currentDir + @"\EurToUSD.html", "uccResultAmount", 120);
                    ss.SpeakAsync(valueEurToUsd);
                    txtContents.Text += valueEurToUsd + Environment.NewLine;
                    break;
                case "get battery status":
                    ss.SpeakAsync(CheckBatteryStatus());
                    txtContents.Text += CheckBatteryStatus() + Environment.NewLine;
                    break;

                default:
                    break;
            }
            //txtContents.Text += e.Result.Text.ToString() + Environment.NewLine;


        }
        private void Form1_Load(object sender, EventArgs e) {

        }

        private void btnStop_Click(object sender, EventArgs e) {
            sre.RecognizeAsyncStop();
            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }

        public static string GetLocalIPAddress() {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList) {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork) {
                    return ip.ToString();
                }
            }
            throw new Exception("No network adapters with an IPv4 address in the system!");
        }
        //To check if you're connected or not:
        public static bool CheckIfConnected() {
            return System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable();
        }
        public string CheckBatteryStatus() {
            //BatteryChargeStatus.Text = SystemInformation.PowerStatus.BatteryChargeStatus.ToString();
            //BatteryFullLifetime.Text = SystemInformation.PowerStatus.BatteryFullLifetime.ToString();
            //BatteryLifePercent.Text = SystemInformation.PowerStatus.BatteryLifePercent.ToString();
            //BatteryLifeRemaining.Text = SystemInformation.PowerStatus.BatteryLifeRemaining.ToString();
            //PowerLineStatus.Text = SystemInformation.PowerStatus.PowerLineStatus.ToString();

            return SystemInformation.PowerStatus.BatteryChargeStatus.ToString();
        }
        private string getCurrentDate() {
            return DateTime.UtcNow.Date.ToString("dd/MM/yyyy");
        }
        public string ReturnCurrency(string URLstring, string filename, string searchKey, int index) {
            string returnValue = "";
            try {

                using (WebClient client = new WebClient()) // WebClient class inherits IDisposable
                {
                
                    client.DownloadFile(URLstring, filename);
                
                    string line;

                    System.IO.StreamReader file =
                    new System.IO.StreamReader(filename);
                    while ((line = file.ReadLine()) != null) {
                        bool found = line.Contains(searchKey);
                        if (found) {
                            returnValue = getCurrencyValue(line, index);
                            //txtContents.Text += returnValue + Environment.NewLine;

                            file.Close();
                            break;
                        }
                    }
                    file.Close();
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "error");
            }

            return returnValue;
        }

        public string getCurrencyValue(string myString, int index) {
            //120 for euro to smth
            //2 for bitcoin
            char[] delimiterChars = { '>','<' };
            string[] words = myString.Split(delimiterChars);

            return words[index];
        }
    }
}
