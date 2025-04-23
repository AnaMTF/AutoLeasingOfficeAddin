using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using WordClientNou.Models;
using System.Data.OleDb;
using System.Net.NetworkInformation;


namespace PowerPointCatalog
{
    public partial class RibbonPowerpoint
    {
        PowerPoint.Application pptApp;

        private void RibbonPowerpoint_Load(object sender, RibbonUIEventArgs e)
        {
            pptApp = Globals.ThisAddIn.Application;
        }

        private void btnOferta_Click_1(object sender, RibbonControlEventArgs e)
        {
            // 1️⃣ – Slide activ
            PowerPoint.Slide curSlide = null;

            if (pptApp.ActiveWindow?.View?.Type == PowerPoint.PpViewType.ppViewNormal)
                curSlide = pptApp.ActiveWindow.View.Slide;
            else if (pptApp.ActiveWindow?.Selection?.SlideRange?.Count > 0)
                curSlide = pptApp.ActiveWindow.Selection.SlideRange[1];

            if (curSlide == null)
            {
                MessageBox.Show("Selectați un slide cu ofertă.");
                return;
            }

            // 2️⃣ – Obține OfferID
            string idText = curSlide.Tags["OfferID"];
            if (string.IsNullOrEmpty(idText) || !int.TryParse(idText, out int ofertaId))
            {
                MessageBox.Show("Slide-ul selectat nu conține o ofertă validă.");
                return;
            }

            // 3️⃣ – Obține ID-ul clientului din Word
            int idClient = StareClient.IdClientCurent;
            if (idClient <= 0)
            {
                MessageBox.Show("Clientul nu a fost setat. Vă rugăm să adăugați un client în Word înainte de a selecta o ofertă.");
                return;
            }

            // 4️⃣ – UPDATE în Access: setează IdClient + StatusOferta
            string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";

            try
            {
                using (var con = new OleDbConnection(connString))
                {
                    con.Open();
                    string sql = "UPDATE Oferte SET IdClient = ?, StatusOferta = ? WHERE IdOferta = ?";
                    using (var cmd = new OleDbCommand(sql, con))
                    {
                        cmd.Parameters.AddWithValue("?", idClient);
                        cmd.Parameters.AddWithValue("?", "Indisponibil");
                        cmd.Parameters.AddWithValue("?", ofertaId);

                        int rows = cmd.ExecuteNonQuery();
                        if (rows > 0)
                            MessageBox.Show("Oferta a fost rezervată cu succes!");
                        else
                            MessageBox.Show("Eroare: Oferta nu a fost găsită în baza de date.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la actualizarea ofertei: " + ex.Message);
            }
        }
    }
}
