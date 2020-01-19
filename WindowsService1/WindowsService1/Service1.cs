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
        private String date;

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
            this.counter = 0;
            

            // On configure un timer pour se déclencher toutes les minutes.  
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 60000; // 60 secondes  
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();


        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            
            if (this.counter<1)
            {
                LectureConfiguration();           
                this.counter += 1;                
            }

            //appel de la récupération des applications controllées toute les minutes au cas où un paramètre est changé au cours de la journée :
            getApplicationControlled(this.ConnectionString);


            // Tracer le nombre d'exécution du traitement
            eventLog1.WriteEntry("Exécution du traitement", EventLogEntryType.Information, eventId++);

            Process[] processes = Process.GetProcesses();

            List<String> processN = new List<string>();
            foreach(Process process in processes)
            {
                if (!processN.Contains(process.ProcessName))
                {
                    processN.Add(process.ProcessName);
                    if (this.appControlled.Contains(process.ProcessName))
                    {
                        //Change le jour si l'utilisateur change de jour. 
                        if (DateTime.Now.DayOfWeek.ToString() != this.date)
                        {
                            changeTps(DateTime.Now.DayOfWeek.ToString(), process.ProcessName);
                            resetTpsArret(process.ProcessName);
                            this.date = DateTime.Now.DayOfWeek.ToString();
                        }
                        
                        try
                        {
                            checkEndBlocage(process.ProcessName);

                            string requete = "SELECT Tps_exe_restant FROM ApplicationControlee ap" +
                                " WHERE ap.Id_app = (SELECT Id_app FROM ApplicationControlable WHERE Nom_app = '" + process.ProcessName + "')";

                            // Initialiser la data source avec le résultat de la requête.
                            int a = GetData(this.ConnectionString, requete);

                            if (a == 1 || a==0)
                            {
                                foreach (Process p in Process.GetProcessesByName(process.ProcessName))
                                {
                                    p.Kill();
                                }
                                setTpsArret(process.ProcessName);
                                updateHistory(process.ProcessName, DateTime.Now.DayOfWeek.ToString());
                            } else if (a > 1)
                            {
                                //Change le temps restant si l'admin change entre temps l'état du jour (passe de jour controllé à jour non controllé)
                                if(isTempsActif(DateTime.Now.DayOfWeek.ToString(), process.ProcessName) == -1)
                                {
                                    changeTps(DateTime.Now.DayOfWeek.ToString(), process.ProcessName);
                                    updateHistory(process.ProcessName, DateTime.Now.DayOfWeek.ToString());
                                } else
                                {
                                    decreaseTpsExe(process.ProcessName);

                                }
                            } else
                            {
                                //test si l'admin change l'activation du controle pour le jour donné, on modifie le tps d'execution :
                                changeTps(DateTime.Now.DayOfWeek.ToString(), process.ProcessName);
                            }


                        }
                        catch { }
                    }
                }
            }
        }

        //Défini la liste des noms des applications controllées en fonction de l'état du controle :
        private void getApplicationControlled(string conn)
        {
            this.appControlled = new List<string>();

            SqlConnection AppSmartConnection = new SqlConnection(conn);
            AppSmartConnection.Open();
            String request = "SELECT Nom_app FROM [dbo].[ApplicationControlable] ac LEFT JOIN ApplicationControlee ap ON ap.Id_app = ac.Id_app WHERE ap.Est_actif = 1 ";
            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(request, AppSmartConnection);


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

        //change le temps d'éxecution restant de l'application en fonction du jour de la semaine
        private void changeTps(string jourAng, string nomApp)
        {
            int tps = isTempsActif(jourAng, nomApp);

            string request = "UPDATE ApplicationControlee SET Tps_exe_restant = " + tps + " WHERE Id_app = (SELECT id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "')";
            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand newcommand = new SqlCommand(request, AppSmartConnection);
            AppSmartConnection.Open();
            newcommand.ExecuteNonQuery();
            AppSmartConnection.Close();
            

        }

        private int isTempsActif(string jourAng, string nomApp)
        {
            String jourFr;
            int tps = -1;
            switch (jourAng)
            {
                case "Monday":
                    jourFr = "Lundi";
                    break;
                case "Tuesday":
                    jourFr = "Mardi";
                    break;
                case "Wednesday":
                    jourFr = "Mercredi";
                    break;
                case "Thursday":
                    jourFr = "Jeudi";
                    break;
                case "Friday":
                    jourFr = "Vendredi";
                    break;
                case "Saturday":
                    jourFr = "Samedi";
                    break;
                default:
                    jourFr = "Dimanche";
                    break;
            }

            string requete = "SELECT " + jourFr + " FROM TempsDefini tps WHERE tps.Id_app = (SELECT Id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "') AND tps." + jourFr + "_actif = 1";
            // Création d'un objet connexion en se basant sur la chaine de connexion.

            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            try
            {
                tps = (int)command.ExecuteScalar();
            }
            catch { }
            AppSmartConnection.Close();
            return tps;
        }

        private void decreaseTpsExe(string processName)
        {
            string requete = "UPDATE ApplicationControlee SET Tps_exe_restant = Tps_exe_restant - 1 WHERE Id_app = (SELECT id_app FROM ApplicationControlable WHERE Nom_app = '" + processName + "')";
            // Création d'un objet connexion en se basant sur la chaine de connexion.
            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            command.ExecuteNonQuery();
            AppSmartConnection.Close();
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

        private void setTpsArret(String name)
        {
            //Met la colonne arret concernant l'application stoppée au dateTime du moment de l'arret (moment actuel):
            if (getTpsArret(name) == new DateTime(1754,1,1,0,0,0))
            {
                string requete = "UPDATE ApplicationControlee SET Tps_limite_atteinte = '" + DateTime.Now + "' WHERE Id_app = (SELECT id_app FROM ApplicationControlable WHERE Nom_app = '" + name + "')";

                // Création d'un objet connexion en se basant sur la chaine de connexion.
                SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

                // Création d'un objet commande basé sur la requête reçue en paramètre.
                SqlCommand command = new SqlCommand(requete, AppSmartConnection);
                AppSmartConnection.Open();
                //update a faire donc on utilise la fonction ExecuteNonQuery();
                command.ExecuteNonQuery();
                AppSmartConnection.Close();
            }
        }

        private DateTime getTpsArret(String nomApp)
        {
            string requete = "SELECT Tps_limite_atteinte FROM ApplicationControlee ap WHERE ap.Id_app = (SELECT Id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "')";
            // Création d'un objet connexion en se basant sur la chaine de connexion.

            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            DateTime dt = (DateTime) command.ExecuteScalar();
            AppSmartConnection.Close();
            return dt;
        }

        private void checkEndBlocage(String nomApp)
        {
            //tps moins 1 pour prendre en compte le timer d'1 minute écoulée.
            int dureeBlocage = getTempsBlocage(nomApp);
            DateTime tpsStop = getTpsArret(nomApp);
            TimeSpan intervalle = DateTime.Now - tpsStop;


            //1440 minutes = 1 journée, permet de ne pas reset le temps si tpsStop = valeur par defaut = 01/01/1754
            if (intervalle.Minutes > dureeBlocage && intervalle.TotalDays < 1)
            {
                changeTps(DateTime.Now.DayOfWeek.ToString(), nomApp);
                resetTpsArret(nomApp);
            }

        }

        private void resetTpsArret(String nomApp)
        {
            //reset du temps_limite
            string requete = "UPDATE ApplicationControlee SET Tps_limite_atteinte = '" + new DateTime(1754, 1, 1, 0, 0, 0) + "' WHERE Id_app = (SELECT id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "')";

            // Création d'un objet connexion en se basant sur la chaine de connexion.
            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            //update a faire donc on utilise la fonction ExecuteNonQuery();
            command.ExecuteNonQuery();
            AppSmartConnection.Close();
        }

        private int getTempsBlocage(String nomApp)
        {
            string requete = "SELECT Duree_blocage FROM TempsDefini tps WHERE tps.Id_app = (SELECT Id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "')";
            // Création d'un objet connexion en se basant sur la chaine de connexion.

            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            int tps = (int)command.ExecuteScalar();
            AppSmartConnection.Close();
            return tps;
        }

        private void updateHistory(String nomApp, String jourAng)
        {
            int tempsMax = isTempsActif(jourAng, nomApp);

            int q = -1;

            //Teste si la date existe dans la BDD :
            string requete = "Select Temps_total FROM Historique WHERE Id_app = (SELECT id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "') AND Date='"+DateTime.Now.ToString("d") + "'";

            // Création d'un objet connexion en se basant sur la chaine de connexion.
            SqlConnection AppSmartConnection = new SqlConnection(this.ConnectionString);

            // Création d'un objet commande basé sur la requête reçue en paramètre.
            SqlCommand command = new SqlCommand(requete, AppSmartConnection);
            AppSmartConnection.Open();
            try
            {
                q=(Int32)command.ExecuteScalar();
            }
            catch { }
            AppSmartConnection.Close();

            //récupération de l'id de l'application :
            requete = "Select Id_app FROM ApplicationControlable WHERE Nom_app = '" + nomApp + "'";
            SqlConnection AppSmartCo = new SqlConnection(this.ConnectionString);
            SqlCommand newcommand = new SqlCommand(requete, AppSmartCo);
            AppSmartCo.Open();
            int id = (int)newcommand.ExecuteScalar();
            AppSmartCo.Close();

            //récupération du temps d'éxécution de l'appli :
            requete = "Select Tps_exe_restant FROM ApplicationControlee WHERE Id_app = " + id;
            SqlCommand commande = new SqlCommand(requete, AppSmartCo);
            AppSmartCo.Open();
            int tpsExeRestant = (int)commande.ExecuteScalar();
            AppSmartCo.Close();

            if (tpsExeRestant == -1)
            {
                tpsExeRestant = tempsMax + 1;
            }

            if (q==-1)
            {
                int tps = tempsMax - tpsExeRestant + 1;
                //ajoute la ligne avec le temps passé :
                requete = "INSERT INTO Historique (Id_app,Date,Temps_total) values (@id,@date,@Temps_Total)";
                SqlCommand command2 = new SqlCommand(requete, AppSmartCo);
                command2.Parameters.AddWithValue("@id", id);
                command2.Parameters.AddWithValue("@date", DateTime.Now.Date.ToString("d"));
                //+1 puisque on arrete quand le tps_exe vaut 1,
                command2.Parameters.AddWithValue("@Temps_total", tps);

                AppSmartCo.Open();
                //update a faire donc on utilise la fonction ExecuteNonQuery();
                command2.ExecuteNonQuery();
                AppSmartCo.Close();
            } else
            {
                int tps = tempsMax - tpsExeRestant - 1;

                //ajout du temps passé :
                requete = "UPDATE Historique SET Temps_total = Temps_total + " + tps  + " WHERE Id_app = "+id+" AND Date = '"+DateTime.Now.Date.ToString("d") + "'";


                // Création d'un objet commande basé sur la requête reçue en paramètre.
                SqlCommand commandUpdate = new SqlCommand(requete, AppSmartConnection);
                AppSmartConnection.Open();
                //update a faire donc on utilise la fonction ExecuteNonQuery();
                commandUpdate.ExecuteNonQuery();
                AppSmartConnection.Close();
            }

            
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

