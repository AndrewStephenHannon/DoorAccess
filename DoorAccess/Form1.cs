using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using DoorAccess.com.mindbodyonline.api;
using System.IO.Ports;
using System.Media;
using System.Reflection;
using System.Text.RegularExpressions;

namespace DoorAccess
{
    public partial class Form1 : Form
    {

        private SerialPort com3 = new SerialPort("COM3", 9600, Parity.None, 8, StopBits.One);
        private SerialPort com4 = new SerialPort("COM4", 9600, Parity.None, 8, StopBits.One);

        private Thread doorControl;
        private Thread timerControl;

        private SoundPlayer verified;
        private SoundPlayer error;

        System.Timers.Timer timer;

        int timeLeft = 4;

        public Form1()
        {
            InitializeComponent();

            numericUpDown1.Value = timeLeft;
            
            doorControl = new Thread(new ThreadStart(DoorController));
            timerControl = new Thread(new ThreadStart(startTimer));

            timer = new System.Timers.Timer();
            timer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
            timer.Interval = 1000;

            doorControl.Start();

            while (!doorControl.IsAlive);
        }

        [STAThread]
        private void DoorController()
        {

            com3.Handshake = Handshake.None;
            com3.RtsEnable = true;

            com3.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived);

            com3.Open();
            com4.Open();

            Application.Run();
        }

        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string indata = sp.ReadExisting();
            textBox2.Text = "Data Received:";

            string result = new string(indata.Where(c => char.IsDigit(c)).ToArray());
            textBox2.Text = result;

            port_checkData(textBox2.Text);
        }

        private void port_checkData(string data)
        {
            if (data.Length > 0)
            {
                string sourcename = "FullSceneAthleticCentre";
                string sourcepassword = "KmmuB5I36T3BWsSmk8ZmzbQsOKQ=";
                int[] siteIDs = { 211250 };

                ClientService cService = new ClientService();

                GetClientsRequest cRequest = new GetClientsRequest();

                cRequest.SourceCredentials = new SourceCredentials();
                cRequest.SourceCredentials.SourceName = sourcename;
                cRequest.SourceCredentials.Password = sourcepassword;
                cRequest.SourceCredentials.SiteIDs = siteIDs;
                cRequest.PageSize = 1000;
                cRequest.CurrentPageIndex = 0;
                cRequest.ClientIDs = new string[] { data };

                GetClientsResult cResult = cService.GetClients(cRequest);
                
                if (cResult.Clients.Length == 1)
                {
                    foreach (Client client in cResult.Clients)
                    {
                        if (client.Status.Equals("Active"))
                        {
                            string result = client.ID + "\t" + client.FirstName + " " + client.LastName;

                            textBox3.Text = result;

                            verified.Play();

                            com4.Write("A");

                            timerControl = new Thread(new ThreadStart(startTimer));
                            timerControl.Start();

                            AddArrivalRequest aRequest = new AddArrivalRequest();
                            aRequest.SourceCredentials = new SourceCredentials();
                            aRequest.SourceCredentials.SourceName = sourcename;
                            aRequest.SourceCredentials.Password = sourcepassword;
                            aRequest.SourceCredentials.SiteIDs = siteIDs;
                            aRequest.ClientID = client.ID;
                            aRequest.LocationID = 1;

                            AddArrivalResult aResult = cService.AddArrival(aRequest);
                        }
                        else
                        {
                            error.Play();
                        }
                    }
                }
                else
                {
                    error.Play();
                }
            }
        }

        private void startTimer()
        {
            bool enabled = false;
            SetBool(enabled);
            timeLeft = (int)numericUpDown1.Value;
            string timeL = "Door is currently open";
            SetText(timeL.ToString());
            timer.Start();
        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {

            if(timeLeft > 1)
            {
                timeLeft = timeLeft - 1;
            }
            else
            {
                timer.Stop();
                com4.Write("a");
                string timeL = "";
                SetText(timeL.ToString());
                timeLeft = (int)numericUpDown1.Value;
                bool enabled = true;
                SetBool(enabled);
                timerControl.Abort();
            }
        }

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            if (this.label2.InvokeRequired)
            {
                SetTextCallback t = new SetTextCallback(SetText);
                this.Invoke(t, new object[] { text });
            }
            else
            {
                this.label2.Text = text;
            }
        }

        delegate void SetBoolCallback(bool value);

        private void SetBool(bool value)
        {
            if(this.numericUpDown1.InvokeRequired)
            {
                SetBoolCallback t = new SetBoolCallback(SetBool);
                this.Invoke(t, new object[] { value });
            }
            else
            {
                this.numericUpDown1.Enabled = value;
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            timeLeft = (int)numericUpDown1.Value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            com4.Write("A");

            timerControl = new Thread(new ThreadStart(startTimer));
            timerControl.Start();
        }

        protected virtual void Form1_OnClosing(object sender, CancelEventArgs e)
        {
            if(doorControl.IsAlive)
                doorControl.Abort();
            if(timerControl.IsAlive)
                timerControl.Abort();

            com3.Close();
            com4.Close();

            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Assembly assembly;
            assembly = Assembly.GetExecutingAssembly();
            verified = new SoundPlayer(assembly.GetManifestResourceStream("DoorAccess.Valid Sound.wav"));
            error = new SoundPlayer(assembly.GetManifestResourceStream("DoorAccess.Error Sound.wav"));
        }
    }
}
