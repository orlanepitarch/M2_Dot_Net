using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        private int eventId = 1;


        public Service1()
        {
            InitializeComponent();

            eventLog1 = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("Test"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "Test", "MonNouveauLog");
            }

            // Définir le nom de la source à utiliser pendant l'écriture dans le journal
            // des événements
            eventLog1.Source = "Test";

            // Définir le nom du journal à utiliser en lecture et en écriture.
            eventLog1.Log = "MonNouveauLog";

        }

        protected override void OnStart(string[] args)
        {
            // Ecrire dans le journal des événements
            eventLog1.WriteEntry("Dans la méthode OnStart");
            // On configure un timer pour se déclencher toutes les minutes.  
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 secondes  
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();


        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            // Tracer le nombre d'exécution du traitement
            eventLog1.WriteEntry("Exécution du traitement", EventLogEntryType.Information, eventId++);

            // TODO: Insérer ici le traitement que doit faire le service.

            Process[] processes = Process.GetProcesses();
            Console.WriteLine("Count: {0}", processes.Length);
            WriteToFile("Nouveau Log");
            foreach(Process process in processes)
            {
                WriteToFile(process.ProcessName);
            }
            List<String> toStop = readToFile();
            foreach (String process in toStop)
            {
                foreach (Process p in Process.GetProcessesByName(process)) {
                    p.Kill();
                }
            }
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        public List<String> readToFile()
        {
            List<String> toStop = new List<String>();
            String path = @"C:\\Users\Orlane\test.txt";
            using (StreamReader sr = new StreamReader(path))
            {
                while (sr.Peek() >= 0)
                {
                    toStop.Add(sr.ReadLine());

                }
            }
            return toStop;
        }


        protected override void OnStop()
        {
            eventLog1.WriteEntry("Dans la méthode OnStop.");
            // TODO: Insérer ici le traitement que doit faire le service
            // Lorsqu'il reçoit une commande Arrêter.
        }

        protected override void OnContinue()
        {
            eventLog1.WriteEntry("Dans la méthode OnContinue.");
        }

        protected override void OnPause()
        {
            eventLog1.WriteEntry("Dans la méthode OnPause.");
        }


    }
}

