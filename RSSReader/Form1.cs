/* Programma scritto da Davide Caminati il 3/8/2014
 * davide.caminati@gmail.com
 * http://caminatidavide.it/
 * 
 * licenza copyleft 
 * http://it.wikipedia.org/wiki/Copyleft#Come_si_applica_il_copyleft
 */


// TODO 
// cleanup del codice


/* Example of c:\configfile.txt  the configuration file
@http://feeds.ilsole24ore.com/c/32276/f/438662/index.rss
@http://feeds.ilsole24ore.com/c/32276/f/566660/index.rss
|COM9
 */ 

// COMANDI
// ESC per uscire
// INVIO per leggere articolo
// slide per cambiare articolo
// button1 per stop
// button2 per pause
// button3 per play

// NOTE
// address.Text = "http://webvoice.tingwo.co/ilsole5642813vox?url=";


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.ServiceModel.Syndication;
using System.Speech.Synthesis;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;
using System.IO.Ports;
using System.Diagnostics;

namespace RSSReader
{
    public partial class Form1 : Form
    {
        string LastUrl = "";
        WebBrowser feedview = new WebBrowser();
        bool onPause = false;
        int actualindex = 0;
        int buttonState1 = 1;
        int buttonState2 = 1;
        int buttonState3 = 1;
        string ComPortName = "";

        WMPLib.WindowsMediaPlayer Player;
        SerialPort mySerialPort ;

        // Initialize a new instance of the SpeechSynthesizer.
        SpeechSynthesizer synth = new SpeechSynthesizer();

        public Form1()
        {
            InitializeComponent();
            ParlaBloccante("caricamento lista articoli, attendere");
            feedview.DocumentCompleted += feedview_DocumentCompleted;
            // Configure the audio output. 
            synth.SetOutputToDefaultAudioDevice();
            ReadConfigFile();
            mySerialPort = new SerialPort(ComPortName);
        }


        private void Read_RSS(string indirizzo)
        {
            //listfeed.Items.Clear(); 
            try
            {
                XmlReader reader = XmlReader.Create(indirizzo);
                SyndicationFeed feed = SyndicationFeed.Load(reader);
                foreach (SyndicationItem s in feed.Items)
                {
                    string[] r = { s.Title.Text, s.Links[0].Uri.ToString() };
                    listfeed.Items.Add(new ListViewItem(r));
                }
                if (listfeed.Items.Count == 0)
                {
                    ParlaBloccante("errore nel caricamento degli articoli del sole 24 ore");
                }
                else
                {
                    //Parla("Lista articoli " + tipo + " caricata");
                }
            }
            catch (Exception q) 
            {
                ParlaBloccante("Errore nella lettura del feed. messaggio:" + q.Message);
            }
        }


        private void Muto()
        {

            synth.SpeakAsyncCancelAll();
        }


        private void Parla(string args)
        {
            Muto();
            synth.SpeakAsyncCancelAll();
            //synth.Speak(args);
            synth.SpeakAsync(args);
        }


        private void ParlaBloccante(string args)
        {
            Muto();
            synth.SpeakAsyncCancelAll();
            synth.Speak(args);
        }


        private string isBorderElement(ListViewItem elemento)
        {
            string terminatoreFrase;
            var r = Enumerable.Empty<ListViewItem>();
            r = this.listfeed.Items.OfType<ListViewItem>();
            var last = r.LastOrDefault();
            var first = r.FirstOrDefault();
            if (listfeed.SelectedItems[0] == last)
            {
                terminatoreFrase = ". ultimo elemento";
            }
            else if (listfeed.SelectedItems[0] == first)
            {
                terminatoreFrase = ". primo elemento";
            }
            else
            {
                terminatoreFrase = "";
            }
            return terminatoreFrase;
        }


        private void listfeed_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listfeed.SelectedItems.Count == 1)
            {
                //stop eventualy other file
                StopFile();

                int selectedIndex = listfeed.SelectedIndices[0];
                try
                {
                    if (feedview.Document != null)
                    {
                        feedview.Document.OpenNew(true);
                        feedview.Document.Write(listfeed.SelectedItems[0].SubItems[1].Text);
                    }
                    else
                    {
                        feedview.DocumentText = listfeed.SelectedItems[0].SubItems[1].Text;
                    }
                    string terminatore = isBorderElement(listfeed.SelectedItems[0]);
                    Parla(listfeed.SelectedItems[0].Text.ToString() + terminatore);
                }
                catch { }
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            feedview.ScriptErrorsSuppressed = true;

            if (listfeed.Items.Count > 0)
            {
                listfeed.Items[0].Selected = true;
                listfeed.Select();
                // Open the serial port
                OpenSerialPort();
            }
            else
            {
                ParlaBloccante("Errore durante caricamento lista articoli. programma bloccato, si consiglia di chiudere il programma");
            }
        }

        
        private void ReadConfigFile()
        {
            try
            {
                string line;

                // Read the file and display it line by line.
                System.IO.StreamReader file = new System.IO.StreamReader(@"c:\configfile.txt");
                while ((line = file.ReadLine()) != null)
                {
                    if (line.StartsWith("@"))
                    {
                        Read_RSS(line.Split('@')[1]);
                    }
                    if (line.StartsWith("|"))
                    {
                        ComPortName = line.Split('|')[1];
                    }
                }

                file.Close();
            }
            catch (Exception ex)
            {
                ParlaBloccante("il file di configurazione non puo' essere aperto causa: " + ex.Message);
            }
        }


        private void OpenSerialPort()
        {
            try
            {
                mySerialPort.BaudRate = 9600;
                mySerialPort.Parity = Parity.None;
                mySerialPort.StopBits = StopBits.One;
                mySerialPort.DataBits = 8;
                mySerialPort.Handshake = Handshake.None;
                mySerialPort.DataReceived += SerialPortDataReceived;
                mySerialPort.Open();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message + ex.StackTrace);
            }
        }


        private void SerialPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            string  buff = sp.ReadLine();
            if (!String.IsNullOrEmpty(buff))
            {
                char[] delimiterChars = { ' ' };
                string[] words = buff.Split(delimiterChars);
                if (words.Count() == 4)
                {
                    int a = Convert.ToInt32(words[0].ToString());
                    SelezionaFeed(a);
                    buttonState1 = Convert.ToInt32(words[1].ToString());
                    buttonState2 = Convert.ToInt32(words[2].ToString());
                    buttonState3 = Convert.ToInt32(words[3].ToString());

                }
            }
        }


        private void SelezionaFeed(int indice)
        {
            //listfeed.Items[indice].Selected = true;
            actualindex = indice;
        }


        private void feedview_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string indirizzo = feedview.Url.ToString();
            if (indirizzo != "about:blank")
            {
                if (!LastUrl.StartsWith("http://webvoice.tingwo.co"))
                {
                    string testo = feedview.DocumentText;
                    string percorsomp3 = RegexLib.FindMp3Path(testo);
                    // download mp3
                    if  (percorsomp3 == "")
                    {
                        if (indirizzo.StartsWith("http://www.ilsole24ore.com"))
                        {
                            if (LastUrl != feedview.Url.ToString())
                            {
                                LastUrl = feedview.Url.ToString();
                                Parla("Lettura articolo"); //in realtà verrà letto al prossimo document_complete
                                feedview.Navigate("http://webvoice.tingwo.co/ilsole5642813vox?url=" + LastUrl);

                            }
                        }
                    }

                    else
                        {
                            // blocca caricamento pagina nel browser
                            feedview.Navigate("about:blank");
                            //start download
                            WebClient webClient = new WebClient();
                            webClient.DownloadFile("http://webvoice.tingwo.co/" + percorsomp3, @"c:\myfile.mp3");

                            //stop eventualy other file
                            StopFile();
                            // open mp3file
                            PlayFile(@"c:\myfile.mp3");

                            //svuota percorsomp3
                            percorsomp3 = "";
                        }
                    }
                }
            }


        private void PlayFile(String url)
        {
            Player = new WMPLib.WindowsMediaPlayer();
            Player.PlayStateChange +=
                new WMPLib._WMPOCXEvents_PlayStateChangeEventHandler(Player_PlayStateChange);
            Player.MediaError +=
                new WMPLib._WMPOCXEvents_MediaErrorEventHandler(Player_MediaError);
            Player.URL = url;
            Player.controls.play();
        }


        private void StopFile()
        {
            try
            {
                Player.controls.stop();
                Player.close();
                onPause = false;
            }
            catch
            { }
        }


        private void Player_PlayStateChange(int NewState)
        {
            if ((WMPLib.WMPPlayState)NewState == WMPLib.WMPPlayState.wmppsStopped)
            {
                // to do 
                //for eache file save the position and ask if continue from this position on load
                string posizione = Player.controls.currentPositionString;
            }
        }


        private void Player_MediaError(object pMediaObject)
        {
            ParlaBloccante("errore nel caricamento del file");
        }


        public class RegexLib
        {
            public static string FindMp3Path(string input)
            {
                string pattern = @"[""].download.+[""]";
                Regex rgx = new Regex(pattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);

                // Find matches.
                MatchCollection matches = rgx.Matches(input);
                string risultato = "";
                if (matches.Count == 1)
                { 
                    risultato = matches[0].ToString();
                    risultato = risultato.Replace ("\"","");
                }
                return risultato;
            }
        }


        private void PauseResumeFile()
        {
            try
            {
                if (onPause)
                {
                    /* Pause the Player. */
                    Player.controls.play();
                    onPause = false;
                }
                else
                {
                    Player.controls.pause();
                    onPause = true;
                }
            }
            catch
            { }
        }


        private void listfeed_KeyUp(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    if ((listfeed.Items.Count > 0) && (listfeed.SelectedItems.Count > 0))
                    {
                        feedview.Navigate(listfeed.SelectedItems[0].SubItems[1].Text);
                    }
                    break;

                case Keys.Escape:
                    StopFile();
                    Console.Beep();
                    Console.Beep();
                    Console.Beep();
                    this.Close();
                    break;

                case Keys.Down:
                    //stop eventualy other file
                    StopFile();
                    break;

                case Keys.Up:
                    //stop eventualy other file
                     StopFile();
                    break;

                case Keys.Space:
                    PauseResumeFile();
                    break;
            }
        }


        // FAST (but not so elegant) solution for cross threading in serial data input
        private void timer1_Tick(object sender, EventArgs e)
        {
            int selectedIndex = listfeed.SelectedIndices[0];
            if (actualindex != selectedIndex)
            {
                listfeed.SelectedItems.Clear();
                listfeed.Items[actualindex].Selected = true;
                listfeed.Items[actualindex].Focused = true;
                listfeed.Select();
            }


            if (buttonState3 == 0)
            {
                // eseguo la lettura dell'articolo
                if ((listfeed.Items.Count > 0) && (listfeed.SelectedItems.Count > 0))
                {
                    feedview.Navigate(listfeed.SelectedItems[0].SubItems[1].Text);
                }
                //resetto la variabile
                buttonState3 = 1;
            }

            if (buttonState2 == 0)
            {
                // metto in pausa la lettura dell'articolo
                PauseResumeFile();
                //resetto la variabile
                buttonState2 = 1;
            }

            if (buttonState1 == 0)
            {
                // interrompo la lettura dell'articolo
                StopFile();
                //resetto la variabile
                buttonState1 = 1;
            }

        }
    }
}
