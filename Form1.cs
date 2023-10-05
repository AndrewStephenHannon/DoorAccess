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

        private SerialPort com3 = new SerialPort("COM3", 9600, Parity.None, 8, StopBits.One);   //used to read in signal from card reader device on door
        private SerialPort com4 = new SerialPort("COM4", 9600, Parity.None, 8, StopBits.One);   //used to send signal to door lock to open it

        private Thread doorControl;     //used to control the state of the door
        private Thread timerControl;    //used to keep track of the time the door is open for

        private SoundPlayer verified;   //used to play sound when card read is accepted
        private SoundPlayer error;      //used to play sound when card read is rejected

        System.Timers.Timer timer;      //timer for door

        int timeLeft = 4;   //amount of time door is open for when card is accepted (in seconds)

        public Form1()
        {
            InitializeComponent();

            numericUpDown1.Value = timeLeft;    //set this field on the form to the default time of 4 seconds. User can change the value from the form by using this field
            
            doorControl = new Thread(new ThreadStart(DoorController));  //initialize the thread for controlling the door
            timerControl = new Thread(new ThreadStart(startTimer));     //initialize the thread for tracking time

            timer = new System.Timers.Timer();  //initialize timer
            timer.Elapsed += new ElapsedEventHandler(OnTimedEvent); //add an alapsed timer event
            timer.Interval = 1000;  //interval for how often the elapsed timer event occurs

            doorControl.Start();        //start the door control thread

            while (!doorControl.IsAlive);       //checks if thread terminated
        }


        //Door Controller thread for setting up communications to and from door
        [STAThread]
        private void DoorController()
        {
            //set up COM3 to receive data from the door lock device
            com3.Handshake = Handshake.None;
            com3.RtsEnable = true;
            com3.DataReceived += new SerialDataReceivedEventHandler(port_DataReceived); //create event handler for when signal/data is received from door lock device

            com3.Open();    //open COM3
            com4.Open();    //open COM4

            Application.Run();
        }

        //method for handling received signal from door lock
        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;     //instantiate serial port to obtain data that was received
            string indata = sp.ReadExisting();      //store received data in string

            //display the data received (the member ID from the card that was scanned) in the Data Received text box for application administrator
            //to see the member ID in case a manual override is needed on the door lock
            string result = new string(indata.Where(c => char.IsDigit(c)).ToArray());
            textBox2.Text = result;

            //call method to check if member ID is in active member database
            port_checkData(textBox2.Text);
        }

        //accesses MindBody account for gym to compare member ID read from door lock device with member IDs of active memberships
        private void port_checkData(string data)
        {
            if (data.Length > 0)
            {
                //access the MindBody account
                //username and password omitted for security purposes
                string sourcename = "*************";
                string sourcepassword = "**************";
                int[] siteIDs = { ****** };

                ClientService cService = new ClientService();           //instantiate a client service using the MindBody API
                GetClientsRequest cRequest = new GetClientsRequest();   //instantiate a request to obtain information from the MindBody database

                //create the request using the login credentials from above
                cRequest.SourceCredentials = new SourceCredentials();
                cRequest.SourceCredentials.SourceName = sourcename;
                cRequest.SourceCredentials.Password = sourcepassword;
                cRequest.SourceCredentials.SiteIDs = siteIDs;
                cRequest.PageSize = 1000;       //set results page size to 1000, this can be adjusted to clients needs
                cRequest.CurrentPageIndex = 0;  //set the page index to the firs tpage
                cRequest.ClientIDs = new string[] { data }; //set the that is being compared for the request (member ID that was read in from door lock device)

                GetClientsResult cResult = cService.GetClients(cRequest);
                
                //if one member is found from the member ID, then check if the member has an active membership
                if (cResult.Clients.Length == 1)
                {
                    //get the member that was found (could also use cResult.Clients[0] rather than for loop since we know at this point there is only one result)
                    foreach (Client client in cResult.Clients)
                    {
                        //check if member has an active membership
                        if (client.Status.Equals("Active"))
                        {
                            //if the membership is active, get the member's information (ID, first and last name)
                            string result = client.ID + "\t" + client.FirstName + " " + client.LastName;

                            textBox3.Text = result; //populate the text box under label "Most Recent Member" to show application owner who scanned and was successfully accept

                            verified.Play();    //play the sound effect that indicates card was accepted

                            com4.Write("A");    //send signal to COM4 to tell the door lock device to unlock

                            //start the open timer that will keep the door unlock for the duration indicated by the timer
                            timerControl = new Thread(new ThreadStart(startTimer));
                            timerControl.Start();

                            //create an arrival request with the MindBody API to create a log entry of the successful card scan
                            AddArrivalRequest aRequest = new AddArrivalRequest();
                            aRequest.SourceCredentials = new SourceCredentials();
                            aRequest.SourceCredentials.SourceName = sourcename;
                            aRequest.SourceCredentials.Password = sourcepassword;
                            aRequest.SourceCredentials.SiteIDs = siteIDs;
                            aRequest.ClientID = client.ID;
                            aRequest.LocationID = 1;

                            AddArrivalResult aResult = cService.AddArrival(aRequest);   //add the arrival to the log
                        }
                        //if member does not have an active membership, play the sound to indicate card was rejected
                        else
                        {
                            error.Play();
                        }
                    }
                }
                //if more than one result was received, there is an issue with two conflicting member IDs and sound to indicate card was rejected will play
                else
                {
                    error.Play();
                }
            }
        }

        //starts the timer countdown for amount of time the door remains unlocked for
        private void startTimer()
        {
            //disable the numericUpDown so that the application user can't alter the time setting while the door is open (timer is running)
            bool enabled = false;
            SetBool(enabled);
            timeLeft = (int)numericUpDown1.Value;

            //set label2 to indicate the door is currently open
            string timeL = "Door is currently open";  
            SetText(timeL.ToString());

            //start timer
            timer.Start();
        }

        //countsdown the timer and resets it when it reaches 0
        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {

            //if more than 1 second left in timer, decrement timer be 1 second (remember this method is invoked once every second when time is active)
            if(timeLeft > 1)
            {
                timeLeft = timeLeft - 1;
            }
            //Once timer reaches zero lock door and reset values
            else
            {
                timer.Stop();       //stop the timer
                com4.Write("a");    //send signal to door lock device to lock the door
                string timeL = "";
                SetText(timeL.ToString());      //reset label2 text to empty string (clears door open status)
                timeLeft = (int)numericUpDown1.Value;   //set time left for next timer countdown to the setting in the numericUpDown box
                bool enabled = true;
                SetBool(enabled);       //set the numericUpDown to enabled to allow user to set the countdown time again
                timerControl.Abort();   //End the timer thread
            }
        }

        delegate void SetTextCallback(string text);

        //Set text value of label2 to the value passed through. Invoke called to see if thread ID
        //calling this method is the same as the creating thread. If so, label2 is successful enabled/disabled
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

        //Set wether the numericUpDown for setting the time the door is open for to enabled or disabled for editting based on the value passed through.
        //Invoke called to see if thread ID calling this method is the same as the creating thread. If so, numericUpDown is successful enabled/disabled
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

        //when the value in numericUpDown is change, set timeLeft to its value
        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            timeLeft = (int)numericUpDown1.Value;
        }

        //When application user clicks the "Open door manually" button, door is opened for a period of time.
        //This is used if there are issues with reading valid member card or to let visitors in.
        private void button1_Click(object sender, EventArgs e)
        {
            com4.Write("A");        //send signal to door lock device to unlock

            //start a timer thread to have the door remain open for the set amount of time
            timerControl = new Thread(new ThreadStart(startTimer));
            timerControl.Start();
        }

        //Method for handling coms and thread when application is closed
        protected virtual void Form1_OnClosing(object sender, CancelEventArgs e)
        {
            //Ends the door control thread
            if(doorControl.IsAlive)
                doorControl.Abort();

            //ends the timer control thread
            if(timerControl.IsAlive)
                timerControl.Abort();

            //closes the coms for sending and receiving signals
            com3.Close();
            com4.Close();

            //exits the application
            Application.Exit();
        }

        //load in the sound files for indicating accepted or rejected card scans
        private void Form1_Load(object sender, EventArgs e)
        {
            Assembly assembly;
            assembly = Assembly.GetExecutingAssembly();
            verified = new SoundPlayer(assembly.GetManifestResourceStream("DoorAccess.Valid Sound.wav"));
            error = new SoundPlayer(assembly.GetManifestResourceStream("DoorAccess.Error Sound.wav"));
        }
    }
}
