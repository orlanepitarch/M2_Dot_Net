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
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        private int eventId = 1;
        private string ConnectionString;
        private List<String> appControlled;
        private int counter;
        private DataTable table;

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

        private void getApplicationControlled(string conn)
        {
            this.table = new DataTable();
            this.appControlled = new List<string>();
            
            SqlConnection AppSmartConnection = new SqlConnection(conn);
            AppSmartConnection.Open();
            String request = "SELECT Nom_app FROM [dbo].[ApplicationControlable]";
            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(request, AppSmartConnection);
            
            SqlDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = command;

            // Remplissage de la DataTable avec le résultat de la requête.
            adapter.Fill(this.table);
           
            SqlDataReader reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    this.appControlled.Add(reader.GetString(0));
                }
            }
            else
            {
                eventLog1.WriteEntry("No rows found.");
            }
            reader.Close();
            AppSmartConnection.Close();
        }

        protected override void OnStart(string[] args)
        {
            // Ecrire dans le journal des événements
            eventLog1.WriteEntry("Dans la méthode OnStart");
            this.counter = 0;

            // On configure un timer pour se déclencher toutes les minutes.  
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 secondes  
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();


        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            eventLog1.WriteEntry("Dans la méthode onTimer");
            if (this.counter<1)
            {
                LectureConfiguration();

                getApplicationControlled(this.ConnectionString);
                
                this.counter += 1;
                
            }

            // Tracer le nombre d'exécution du traitement
            eventLog1.WriteEntry("Exécution du traitement", EventLogEntryType.Information, eventId++);

            Process[] processes = Process.GetProcesses();
            WriteToFile("Nouveau Log");


            List<String> processN = new List<string>();
            foreach(Process process in processes)
            {
                if (!processN.Contains(process.ProcessName))
                {
                    processN.Add(process.ProcessName);
                    if (this.appControlled.Contains(process.ProcessName))
                    {
                        eventLog1.WriteEntry(process.ProcessName, EventLogEntryType.Information, eventId++);
                        try
                        {
                            string requete = "SELECT Tps_exe_restant FROM ApplicationControlee ap" +
                                " WHERE ap.Id_app = (SELECT Id_app FROM ApplicationControlable WHERE Nom_app = '" + process.ProcessName + "')";

                            eventLog1.WriteEntry(requete, EventLogEntryType.Information, eventId++);
                            // Initialiser la data source avec le résultat de la requête.
                            int a = GetData(this.ConnectionString, requete);
                            eventLog1.WriteEntry(a.ToString(), EventLogEntryType.Information, eventId++);

                            if (a == 1)
                            {
                                foreach (Process p in Process.GetProcessesByName(process.ProcessName))
                                {
                                    p.Kill();
                                }
                            } else if (a > 1)
                            {
                                decreaseTpsExe(process.ProcessName);
                            }


                        }
                        catch { }
                    }
                }
            }
        }

        private void decreaseTpsExe(string processName)
        {
            string requete = "UPDATE ApplicationControlee SET Tps_exe_restant = Tps_exe_restant - 1 WHERE Id_app = (SELECT id_app FROM ApplicationControlable WHERE Nom_app = '"+processName+"')";
            // Création d'un objet connexion en se basant sur la chaine de connexion.
            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            command.ExecuteNonQuery();
            AppSmartConnection.Close();
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

        private void LectureConfiguration()
        {
            // Lecture de la dernière chaine de connexion utilisée
            ConnectionStringSettings oCnxCfg = ConfigurationManager.ConnectionStrings["DerniereChaine"];

            if (oCnxCfg != null)
            {
                this.ConnectionString = oCnxCfg.ConnectionString;
            }
            else
            {
                // Recherche de la chaine de connexion partielle à la base de données dans le App.config
                oCnxCfg = ConfigurationManager.ConnectionStrings["ConnexionAppSmartLocker"];

                if (oCnxCfg != null)
                {
                    SqlConnectionStringBuilder oCnxBldr = new SqlConnectionStringBuilder(oCnxCfg.ConnectionString);

                    if (string.IsNullOrEmpty(oCnxBldr.InitialCatalog))
                    {
                        oCnxBldr.InitialCatalog = "AppSmartLocker";
                    }

                    this.ConnectionString = oCnxBldr.ConnectionString;
                }
            }

        }
        private int GetData(string conn, string sqlCommand)
        {
            int q = -1;
            // Création d'un objet connexion en se basant sur la chaine de connexion.
            SqlConnection AppSmartConnection = new SqlConnection(conn);
         

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(sqlCommand, AppSmartConnection);
            AppSmartConnection.Open();
            try
            {
                q = (Int32)command.ExecuteScalar();
            }
            catch { }

            AppSmartConnection.Close();
            return q;

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

        private void eventLog1_EntryWritten(object sender, EntryWrittenEventArgs e)
        {

        }
    }
}

