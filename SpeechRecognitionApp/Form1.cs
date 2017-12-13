using System;
using System.Collections;
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

namespace SpeechRecognitionApp {

    public partial class Form1 : Form {

        Dictionary<string, int> CitiesTimeZonesDicitonary = new Dictionary<string, int>() {
            {"Bucharest", 0 },
            {"Dublin", -1 },
            {"London", -2 },
            {"Barcelona", -1 },
            {"Oslo", -1 },
            {"Paris", -1 },
            {"New York", -9 },
            {"Toronto", -10 },
            {"Moscow", 1 },
            {"Dubai", 2 },
            {"Singapore", 6 },
            {"Honk kong", 6 },
            {"Melbourne", 9 },
            {"Tokyo", 7 },
            {"Jakata", 5 },
        };

        string currentDir = System.AppDomain.CurrentDomain.BaseDirectory;
        ArrayList teams;
        ArrayList scores;
        string country;
        Process myProcess;

        //Form Declaration
        SpeechSynthesizer ss = new SpeechSynthesizer();
        PromptBuilder pb = new PromptBuilder();
        SpeechRecognitionEngine sre = new SpeechRecognitionEngine();
        Choices clist = new Choices(new string[] {
                                    "hello",
                                    "how are you",
                                    "thank you",
                                    "close",
                                    "what is the current time",
                                    "what is the current time in Bucharest",
                                    "what is the current time in London",
                                    "what is the current time in New York",
                                    "what is the current time in Paris",
                                    "what is the current time in Barcelona",
                                    "what is the current time in Oslo",
                                    "what is the current time in Dublin",
                                    "what is the current time in Toronto",
                                    "what is the current time in Moscow",
                                    "what is the current time in Dubai",
                                    "what is the current time in Singapore",
                                    "what is the current time in Honk kong",
                                    "what is the current time in Melbourne",
                                    "what is the current time in Tokyo",
                                    "what is the current time in Jakarta",
                                    "open chrome",
                                    "close chrome",
                                    "open notepad",
                                    "close notepad",
                                    "open windows media player",
                                    "close windows media player",
                                    "open paint",
                                    "close paint",
                                    "get current date",
                                    "get battery status",
                                    "get IP address",
                                    "check Internet connection",
                                    "play youtube video number one", "play youtube video number two", "play youtube video number three",
                                    "get bitcoin value",
                                    "tell currency-exchange", "tell currency-exchange to Dollar",
                                    "get today football results in england", "get today football results in france",
                                    "get today football results in spain", "get today football results in germany"});

        public Form1() {

            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {

            btnStart.Enabled = false;
            btnStop.Enabled = true;

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
            txtContents.Text += e.Result.Text.ToString() + Environment.NewLine;

            switch (e.Result.Text.ToString()) {
                case "hello":
                    ss.SpeakAsync("Hello");
                    break;
                case "how are you":
                    ss.SpeakAsync("I am good, thank you");
                    break;
                case "thank you":
                    ss.SpeakAsync("my pleasure");
                    break;
                case "close":
                    Application.Exit();
                    break;
                case "what is the current time":
                    ss.SpeakAsync("current time is " + DateTime.Now.ToLongTimeString());
                    txtContents.Text += DateTime.Now.ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Bucharest":
                    ss.SpeakAsync("current time in Bucharest is " + getDateTime("Bucharest").ToLongTimeString());
                    txtContents.Text += getDateTime("Bucharest").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in London":
                    ss.SpeakAsync("current time in London is " + getDateTime("London").ToLongTimeString());
                    txtContents.Text += getDateTime("London").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Oslo":
                    ss.SpeakAsync("current time in Oslo is " + getDateTime("Oslo").ToLongTimeString());
                    txtContents.Text += getDateTime("Oslo").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Barcelona":
                    ss.SpeakAsync("current time in Barcelona is " + getDateTime("Barcelona").ToLongTimeString());
                    txtContents.Text += getDateTime("Barcelona").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in New York":
                    ss.SpeakAsync("current time in New York is " + getDateTime("New York").ToLongTimeString());
                    txtContents.Text += getDateTime("New York").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Toronto":
                    ss.SpeakAsync("current time in Toronto is " + getDateTime("Toronto").ToLongTimeString());
                    txtContents.Text += getDateTime("Toronto").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Moscow":
                    ss.SpeakAsync("current time in Moscow is " + getDateTime("Moscow").ToLongTimeString());
                    txtContents.Text += getDateTime("Moscow").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Dubai":
                    ss.SpeakAsync("current time in Dubai is " + getDateTime("Dubai").ToLongTimeString());
                    txtContents.Text += getDateTime("Dubai").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Singapore":
                    ss.SpeakAsync("current time in Singapore is " + getDateTime("Singapore").ToLongTimeString());
                    txtContents.Text += getDateTime("Singapore").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Honk kong":
                    ss.SpeakAsync("current time in Honk kong is " + getDateTime("Honk kong").ToLongTimeString());
                    txtContents.Text += getDateTime("Honk kong").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Melbourne":
                    ss.SpeakAsync("current time in Melbourne is " + getDateTime("Melbourne").ToLongTimeString());
                    txtContents.Text += getDateTime("Melbourne").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Tokyo":
                    ss.SpeakAsync("current time in Tokyo is " + getDateTime("Tokyo").ToLongTimeString());
                    txtContents.Text += getDateTime("Tokyo").ToLongTimeString() + Environment.NewLine;
                    break;
                case "what is the current time in Jakarta":
                    ss.SpeakAsync("current time in Jakarta is " + getDateTime("Jakarta").ToLongTimeString());
                    txtContents.Text += getDateTime("Jakarta").ToLongTimeString() + Environment.NewLine;
                    break;
                case "open chrome":
                    myProcess = Process.Start("chrome", "https://www.google.com");
                    break;
                case "close chrome":
                    myProcess.CloseMainWindow();
                    myProcess.Close();
                    break;
                case "open notepad":
                    myProcess = Process.Start("notepad.exe");
                    break;
                case "close notepad":
                    myProcess.CloseMainWindow();
                    myProcess.Close();
                    break;
                case "open windows media player":
                    myProcess = Process.Start("wmplayer.exe");
                    break;
                case "close windows media player":
                    myProcess.CloseMainWindow();
                    myProcess.Close();
                    break;
                case "open paint":
                    myProcess = Process.Start("mspaint.exe");
                    break;
                case "close paint":
                    myProcess.CloseMainWindow();
                    myProcess.Close();
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
                    txtContents.Text += "Bitcoin : ";
                    string bitcoinValue = ReturnCurrency("https://coinmarketcap.com/currencies/bitcoin/", currentDir + @"\bitcoinFile.html", "quote_price", 4);
                    ss.SpeakAsync(bitcoinValue);
                    txtContents.Text += bitcoinValue + Environment.NewLine;
                    break;
                case "tell currency-exchange":
                    txtContents.Text += "1 Eur To Ron : ";
                    string valueEurtoRon = ReturnCurrency("http://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=RON", currentDir + @"\EurToRon.html", "uccResultAmount", 216);
                    ss.SpeakAsync(valueEurtoRon);
                    txtContents.Text += valueEurtoRon + Environment.NewLine;
                    break;
                case "tell currency-exchange to Dollar":
                    txtContents.Text += "1 Eur To USD : ";
                    string valueEurToUsd = ReturnCurrency("http://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=USD", currentDir + @"\EurToUSD.html", "uccResultAmount", 216);
                    ss.SpeakAsync(valueEurToUsd);
                    txtContents.Text += valueEurToUsd + Environment.NewLine;
                    break;
                case "get battery status":
                    ss.SpeakAsync(CheckBatteryStatus());
                    txtContents.Text += CheckBatteryStatus() + Environment.NewLine;
                    break;
                case "get today football results in england":
                    country = e.Result.Text.ToString().Split(' ').Last();
                    getTodayFootballResults(country);
                    break;
                case "get today football results in france":
                    country = e.Result.Text.ToString().Split(' ').Last();
                    getTodayFootballResults(country);
                    break;
                case "get today football results in spain":
                    country = e.Result.Text.ToString().Split(' ').Last();
                    getTodayFootballResults(country);
                    break;
                case "get today football results in germany":
                    country = e.Result.Text.ToString().Split(' ').Last();
                    getTodayFootballResults(country);
                    break;
                default:
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e) {
        }

        private void btnStop_Click(object sender, EventArgs e) {
            sre.RecognizeAsyncStop();
            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }

        public DateTime getDateTime(string cityName) {
            int value = 0;
            if (CitiesTimeZonesDicitonary.ContainsKey(cityName)) {
                value = CitiesTimeZonesDicitonary[cityName];
            }
            return DateTime.Now.AddHours(value);
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

        public static bool CheckIfConnected() {
            try
            {
                using (var client = new WebClient())
                {
                    using (client.OpenRead("http://clients3.google.com/generate_204"))
                    {
                        return true;
                    }
                }
            }
            catch
            {
               return false;
            }
        }

        public string CheckBatteryStatus() {
            return SystemInformation.PowerStatus.BatteryChargeStatus.ToString();
        }

        public string getCurrentDate() {
            return DateTime.UtcNow.Date.ToString("MM/dd/yyyy");
        }

        public string getCurrencyValue(string myString, int index) {
            char[] delimiterChars = { '>', '<', '"' };
            string[] words = myString.Split(delimiterChars);

            return words[index];
        }

        public string ReturnCurrency(string URLstring, string filename, string searchKey, int index) {
            string returnValue = "";
            try {
                using (WebClient client = new WebClient())
                {
                   client.DownloadFile(URLstring, filename);

                    string line;
                    System.IO.StreamReader file =
                    new System.IO.StreamReader(filename);
                    while ((line = file.ReadLine()) != null) {
                        bool found = line.Contains(searchKey);
                        if (found) {
                            returnValue = getCurrencyValue(line, index);
                            break;
                        }
                    }
                    file.Close();
                }
            } catch (Exception ex) {
                MessageBox.Show(ex.Message, "error");
            }

            return returnValue;
        }

        public string parseData(string myString, int index) {
            char[] delimiterChars = { '>', '<' };
            string[] words = myString.Split(delimiterChars);

            return words[index];
        }

        public void getTodayFootballResults(string country) {
            string URLstring = "http://www.oddstake.com/scores/?" + country;
            string filename =  currentDir + @"\Football.html";
            string searchKey;

            teams = new ArrayList();
            scores = new ArrayList();

            switch (country)
            {
                case "england":
                    txtContents.Text += "Premier League :" + Environment.NewLine;
                    searchKey = "http://www.oddstake.com/scorematch/premier-league/";
                    break;
                case "france":
                    txtContents.Text += "Ligue 1 :" + Environment.NewLine;
                    searchKey = "http://www.oddstake.com/scorematch/ligue-1/";
                    break;
                case "spain":
                    txtContents.Text += "La Liga :" + Environment.NewLine;
                    searchKey = "http://www.oddstake.com/scorematch/laliga/";
                    break;
                case "germany":
                    txtContents.Text += "Bundesliga :" + Environment.NewLine;
                    searchKey = "http://www.oddstake.com/scorematch/bundesliga/";
                    break;

                default:
                    searchKey = "";
                    break;
            }

            try
            {
                using (WebClient client = new WebClient())
                {

                    client.DownloadFile(URLstring, filename);

                    string line;
                    string value;
                    System.IO.StreamReader file =
                    new System.IO.StreamReader(filename);
                    while ((line = file.ReadLine()) != null)
                    {
                        bool foundTeam = line.Contains(searchKey);
                        bool foundAwayScore = line.Contains("sbsa");
                        bool foundHomeScore = line.Contains("sbsh");

                        if (foundTeam)
                        {
                            value = parseData(line, 2);
                            teams.Add(value);
                        } else if (foundAwayScore)
                        {
                            value = parseData(line, 2);
                            scores.Add(value);
                        } else if (foundHomeScore)
                        {
                            value = parseData(line, 2);
                            scores.Add(value);
                        }
                    }
                    file.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "error");
            }

            for (int i = 0; i < teams.Count - 1; i += 2)
                txtContents.Text += " * " + teams[i] + " " + scores[i] + " : " + scores[i + 1] + " " + teams[i + 1] + Environment.NewLine;
        }
    }

}
