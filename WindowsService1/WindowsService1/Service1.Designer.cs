namespace WindowsService1
{
    partial class Service1
    {
        /// <summary> 
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.eventLog1 = new System.Diagnostics.EventLog();

            ((System.ComponentModel.ISupportInitialize)(this.eventLog1)).BeginInit();
            // 
            // MonPremierServiceWindows
            // 
            this.ServiceName = "DotNETService1";
            ((System.ComponentModel.ISupportInitialize)(this.eventLog1)).EndInit();

        }



        #region Code généré par le Concepteur de composants

        /// <summary> 
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        

        #endregion

        private System.Diagnostics.EventLog eventLog1;
    }
}
