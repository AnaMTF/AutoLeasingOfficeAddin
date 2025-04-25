using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Data.OleDb;

namespace PowerPointCatalog
{
    public partial class RibbonPowerpoint
    {
        private PowerPoint.Application pptApp;

        // one place to maintain the connection string
        private const string ConnString =
            @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";

        private void RibbonPowerpoint_Load(object sender, RibbonUIEventArgs e)
        {
            pptApp = Globals.ThisAddIn.Application;
        }

        private void btnOferta_Click(object sender, RibbonControlEventArgs e)
        {
            /* ───────────────────────────────────────────────────────────────
               1. slide curent
            ─────────────────────────────────────────────────────────────── */
            PowerPoint.Slide curSlide = null;
            var wnd = pptApp.ActiveWindow;

            if (wnd?.View?.Type == PowerPoint.PpViewType.ppViewNormal)
                curSlide = wnd.View.Slide;
            else if (wnd?.Selection?.SlideRange?.Count > 0)
                curSlide = wnd.Selection.SlideRange[1];

            if (curSlide == null)
            {
                MessageBox.Show("Selectați un slide cu ofertă.");
                return;
            }

            /* ───────────────────────────────────────────────────────────────
               2. OfferID din tag
            ─────────────────────────────────────────────────────────────── */
            string idText = curSlide.Tags["OfferID"];
            if (string.IsNullOrEmpty(idText) || !int.TryParse(idText, out int ofertaId))
            {
                MessageBox.Show("Slide-ul selectat nu conține o ofertă validă.");
                return;
            }

            /* ───────────────────────────────────────────────────────────────
               3. Clientul curent din tabelul DateActuale
            ─────────────────────────────────────────────────────────────── */
            int idClient = GetCurrentClientId();
            if (idClient <= 0)
            {
                MessageBox.Show("Niciun client activ. Adăugați un client în Word înainte de a selecta o ofertă.");
                return;
            }

            /* ───────────────────────────────────────────────────────────────
               4. Rezervăm oferta și marcăm oferta curentă
            ─────────────────────────────────────────────────────────────── */
            try
            {
                var con = new OleDbConnection(ConnString);
                con.Open();

                const string sql =
                    "UPDATE Oferte SET IdClient = ?, StatusOferta = ? WHERE IdOferta = ?";

                var cmd = new OleDbCommand(sql, con);
                cmd.Parameters.AddWithValue("?", idClient);
                cmd.Parameters.AddWithValue("?", "Indisponibil");
                cmd.Parameters.AddWithValue("?", ofertaId);

                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    SetCurrentOfferId(ofertaId);       // păstrăm oferta curentă (opţional)
                    MessageBox.Show("Oferta a fost rezervată cu succes!");
                }
                else
                    MessageBox.Show("Eroare: Oferta nu a fost găsită în baza de date.");

                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Eroare la actualizarea ofertei: " + ex.Message);
            }
        }

        /*──────────────────────── helper-e context ───────────────────────*/

        private int GetCurrentClientId()
        {
            var con = new OleDbConnection(ConnString);

            con.Open();
            var cmd = new OleDbCommand(
                "SELECT IdClientActual FROM DateActuale WHERE ID = 1", con);
            object val = cmd.ExecuteScalar();
            con.Close();

            return val == DBNull.Value ? -1 : Convert.ToInt32(val);
        }

        private void SetCurrentOfferId(int ofertaId)
        {
            var con = new OleDbConnection(ConnString);
            con.Open();
            var cmd = new OleDbCommand(
                "UPDATE DateActuale SET IdOfertaActuala = ? WHERE ID = 1", con);
            cmd.Parameters.AddWithValue("?", ofertaId);
            cmd.ExecuteNonQuery();
            con.Close();
        }
    }
}
