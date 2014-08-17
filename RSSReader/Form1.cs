/* Programma scritto da Davide Caminati il 3/8/2014
 * davide.caminati@gmail.com
 * http://caminatidavide.it/
 */

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

namespace RSSReader
{
    public partial class Form1 : Form
    {

        string LastUrl = "";
        WebBrowser feedview = new WebBrowser();
        string address;
        string tipoRSS;
        bool onPause = false;

        WMPLib.WindowsMediaPlayer Player;


        // Initialize a new instance of the SpeechSynthesizer.
        SpeechSynthesizer synth = new SpeechSynthesizer();

        public Form1()
        {
            InitializeComponent();
            feedview.DocumentCompleted += feedview_DocumentCompleted;
            // Configure the audio output. 
            synth.SetOutputToDefaultAudioDevice();
        }

        private void Read_RSS(string indirizzo,string tipo)
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
                    Parla("errore nel caricamento degli articoli del sole 24 ore");
                }
                else
                {
                    Parla("Lista articoli " + tipo + " caricata");
                }

            }
            catch (Exception q) 
            {
                MessageBox.Show("Errore nella lettura del feed!Info: \n" + q.Message, "Errore");
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
            Parla("caricamento lista articoli, attendere");

            //address.Text = "http://webvoice.tingwo.co/ilsole5642813vox?url=";


            //caricamento indirizzo rss
            //Prima pagina 
            tipoRSS = "Prima pagina";
            address = "http://feeds.ilsole24ore.com/c/32276/f/438662/index.rss";
            Read_RSS(address, tipoRSS);

            //Notizie 
            /*
            tipoRSS = "Notizie";
            address = "http://feeds.ilsole24ore.com/c/32276/f/566660/index.rss";
            Read_RSS(address, tipoRSS);
            */

            /*Notizie 
            tipoRSS = "Giornale Radio";
            address = "http://www.ilsole24ore.com/rss/Radio24_gr24.xml";
            Read_RSS(address, tipoRSS);
            */
            feedview.ScriptErrorsSuppressed = true;

            if (listfeed.Items.Count > 0)
            {
                listfeed.Items[0].Selected = true;
                listfeed.Select();
            }
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
            Parla("errore nel caricamento del file");
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
    }
}
