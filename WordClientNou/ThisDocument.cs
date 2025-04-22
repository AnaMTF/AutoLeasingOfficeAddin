using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using WordClientNou.Models;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Drawing;
using System.IO;

namespace WordClientNou
{
    public partial class ThisDocument
    {
        Button btnAdaugareClient;
        Button btnPowerpoint;
        
        string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            btnAdaugareClient = new Button();
            btnAdaugareClient.Text = "Adauga Client";
            btnAdaugareClient.Click += new EventHandler(btnAdaugareClient_Click);

            btnPowerpoint = new Button();
            btnPowerpoint.Text = "Deschide Catalogul cu Oferte";
            btnPowerpoint.Click += new EventHandler(btnPowerpoint_Click);

            Globals.ThisDocument.ActionsPane.Controls.Add(btnAdaugareClient);
            Globals.ThisDocument.ActionsPane.Controls.Add(btnPowerpoint);
        }

        private void btnPowerpoint_Click(object sender, EventArgs e)
        {
            
            PowerPoint.Application pptApp;
            try
            {
                pptApp = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
            }
            catch (COMException)
            {
                pptApp = new PowerPoint.Application { Visible = Office.MsoTriState.msoTrue };
            }

            
            PowerPoint.Presentation pres =
                pptApp.Presentations.Add(Office.MsoTriState.msoTrue);
            int nrSlides = pres.Slides.Count;
            var slide = pres.Slides.Add(nrSlides+1,
                         PowerPoint.PpSlideLayout.ppLayoutTitleOnly);

            var titlu = slide.Shapes.Title;
            titlu.TextFrame.TextRange.Text = "Catalog de Oferte";
            titlu.TextFrame.TextRange.Font.Size = 32;
            titlu.TextFrame.TextRange.Font.Color.RGB =
                    ColorTranslator.ToOle(Color.PaleVioletRed);
            titlu.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            titlu.TextFrame.TextRange.ParagraphFormat.Alignment =
                    PowerPoint.PpParagraphAlignment.ppAlignCenter;
            titlu.Fill.Visible = Office.MsoTriState.msoTrue;
            titlu.Fill.Solid();
            titlu.Fill.ForeColor.RGB =
                    ColorTranslator.ToOle(Color.FromArgb(173, 216, 230));

            // ----------------------------------------------------------
            // 2. TEXTBOX  – mai jos & mai lat
            // coordonate: Left, Top, Width, Height
            float left = 60f;
            float top = 240f;      
            float width = 820f;      
            float height = 80f;

            var txtBox = slide.Shapes.AddTextbox(
                Orientation: Office.MsoTextOrientation.msoTextOrientationHorizontal,
                Left: left, Top: top, Width: width, Height: height);

            var rng = txtBox.TextFrame.TextRange;
            rng.Text =
            "Acesta este catalogul nostru cu maşinile disponibile pentru închiriere.\n" +
            "Pentru a selecta oferta dorită, asiguraţi‑vă că slide‑ul aferent este selectat, "
            + "apoi în tab‑ul Add‑ins apăsaţi Selectează Oferta.";
            rng.Font.Size = 20;
            rng.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
            rng.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
            txtBox.Fill.Visible = Office.MsoTriState.msoTrue;
            txtBox.Fill.Solid();
            txtBox.Fill.ForeColor.RGB =
                    ColorTranslator.ToOle(Color.FromArgb(173, 216, 230));

            string imgBgPath = @"C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\ImaginiDB\fundalPPT.png";
            slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
            slide.Background.Fill.UserPicture(imgBgPath);

            List<Oferta> oferte = GetOferte(connString);
            foreach (var oferta in oferte)
                AddOfertaSlide(pres, oferta, Globals.ThisDocument.Path);
        }

        private static List<Oferta> GetOferte(string conn)
        {
            var list = new List<Oferta>();
            string sql = "SELECT * FROM Oferte WHERE StatusOferta='Disponibil' ORDER BY Marca, NumeMasina";

            using (var con = new OleDbConnection(conn))
            using (var cmd = new OleDbCommand(sql, con))
            {
                con.Open();
                using (var r = cmd.ExecuteReader())
                {
                    int ixPoza = r.GetOrdinal("Poza");

                    while (r.Read())
                    {
                        var o = new Oferta
                        {
                            IdOferta = r.GetInt32(r.GetOrdinal("IdOferta")),
                            NumeMasina = r.GetString(r.GetOrdinal("NumeMasina")),
                            Marca = r.GetString(r.GetOrdinal("Marca")),
                            TipCarburant = r.GetString(r.GetOrdinal("TipCarburant")),
                            DurataMinima = r.GetInt32(r.GetOrdinal("DurataMinima")),
                            PretZi = r.GetDecimal(r.GetOrdinal("PretZi")),
                            Kilometraj = r.GetInt32(r.GetOrdinal("Kilometraj")),
                            NrLocuri = r.GetInt32(r.GetOrdinal("NrLocuri")),
                            StatusOferta = r.GetString(r.GetOrdinal("StatusOferta")),
                            Poza = ExtractFirstAttachment(r.GetValue(ixPoza))
                        };
                        list.Add(o);
                    }
                }
            }
            return list;
        }

        // helper pt. Attachment → byte[]
        private static byte[] ExtractFirstAttachment(object raw)
        {
            if (raw is byte[] single) return single;
            if (raw is Array arr && arr.Length > 0 && arr.GetValue(0) is byte[] first)
                return first;
            return null;
        }

        // =======================================================
        //  construieşte un slide dintr‑o ofertă
        private static void AddOfertaSlide(PowerPoint.Presentation pres, Oferta of, string documentFolder)
        {
            var sld = pres.Slides.Add(pres.Slides.Count + 1,
                        PowerPoint.PpSlideLayout.ppLayoutBlank);

            // 0️⃣ – ID ofertă
            sld.Tags.Add("OfferID", of.IdOferta.ToString());

            // 1️⃣ – FOTOGRAFIA MAȘINII (dacă există)
            if (of.Poza != null && of.Poza.Length > 0)
            {
                // fișier TMP unic (ex.  C:\...\Temp\718c2e2f.jpg)
                

                string imgPath = System.IO.Path.GetTempFileName();

                File.WriteAllBytes(imgPath, of.Poza);

                // inserează ca Shape (poți modifica poziția și dimensiunea)
                sld.Shapes.AddPicture(imgPath,
                    Office.MsoTriState.msoFalse,   // LinkToFile = FALSE
                    Office.MsoTriState.msoTrue,    // SaveWithDocument = TRUE
                    60, 270, 300, 200);            // Left, Top, Width, Height

                File.Delete(imgPath);              // imaginea e deja încorporată
            }

            // 2️⃣ – Titlu (marca + model)
            var titluBox = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                100, 30, 700, 60);

            var rngTitlu = titluBox.TextFrame.TextRange;
            rngTitlu.Text = $"{of.Marca} {of.NumeMasina}";
            rngTitlu.Font.Size = 42;
            rngTitlu.Font.Bold = Office.MsoTriState.msoTrue;
            rngTitlu.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            rngTitlu.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignCenter;

            titluBox.Fill.Visible = Office.MsoTriState.msoTrue;
            titluBox.Fill.Solid();
            titluBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);

            // 3️⃣ – Detalii
            var detalii = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                80, 120, 760, 120);

            detalii.Fill.Visible = Office.MsoTriState.msoTrue;
            detalii.Fill.Solid();
            detalii.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(208, 240, 255));

            var rngDet = detalii.TextFrame.TextRange;
            rngDet.Text =
                $"Durată minimă: {of.DurataMinima} zile | Tarif: {of.PretZi:C}/zi\n" +
                $"Carburant: {of.TipCarburant} | Locuri: {of.NrLocuri} | " +
                $"Km incluși/zi: {of.Kilometraj}";
            rngDet.Font.Size = 20;
            rngDet.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
            rngDet.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;

            // 4️⃣ – Preț (text roz)
            var pretBox = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                650, 80, 200, 60);
            pretBox.TextFrame.TextRange.Text = of.PretZi.ToString("0");
            pretBox.TextFrame.TextRange.Font.Size = 56;
            pretBox.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.HotPink);
            pretBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignCenter;

            var leiBox = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                680, 140, 140, 40);
            leiBox.TextFrame.TextRange.Text = "Lei";
            leiBox.TextFrame.TextRange.Font.Size = 28;
            leiBox.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.HotPink);
            leiBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignCenter;
        }
        private void btnAdaugareClient_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(rtcNume.Text) || string.IsNullOrEmpty(rtcPrenume.Text) || string.IsNullOrEmpty(rtcTelefon.Text)
                || string.IsNullOrEmpty(rtcJudet.Text) || string.IsNullOrEmpty(rtcCNP.Text) || string.IsNullOrEmpty(rtcEmail.Text)
                || string.IsNullOrEmpty(rtcAdresa.Text) || string.IsNullOrEmpty(rtcLocSec.Text) || string.IsNullOrEmpty(dlcVenit.Text))
            {
                MessageBox.Show("Vă rugăm să completați toate câmpurile.");
                return;
            }
            else if (rtcCNP.Text.Length != 13)
            {
                MessageBox.Show("CNP-ul trebuie să aibă 13 caractere.");
                return;
            }
            else if (rtcTelefon.Text.Length != 10)
            {
                MessageBox.Show("Numărul de telefon trebuie să aibă 10 caractere.");
                return;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(rtcEmail.Text, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                MessageBox.Show("Adresa de email nu este validă.");
                return;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(rtcCNP.Text, @"^\d+$"))
            {
                MessageBox.Show("CNP-ul trebuie să conțină doar cifre.");
                return;
            }
            else
            {
                OleDbConnection conn = new OleDbConnection(connString);
                conn.Open();
                Client client = new Client(rtcNume.Text, rtcPrenume.Text, rtcTelefon.Text, rtcJudet.Text, rtcCNP.Text, rtcEmail.Text, rtcAdresa.Text, rtcLocSec.Text, dlcVenit.Text);
                string maxIdQuery = "SELECT MAX(Id) FROM Clienti";
                OleDbCommand maxIdCommand = new OleDbCommand(maxIdQuery, conn);
                object result = maxIdCommand.ExecuteScalar();
                int maxId = (result == DBNull.Value) ? 0 : System.Convert.ToInt32(result);
                client.Id = maxId + 1;
                string insertQuery = "INSERT INTO Clienti (Nume, Prenume, Telefon, CNP, Email, Judet, LocSec, Adresa, Venit, DataCreare) " +
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                OleDbCommand insertCommand = new OleDbCommand(insertQuery, conn);
                insertCommand.Parameters.AddWithValue("?", client.Nume);
                insertCommand.Parameters.AddWithValue("?", client.Prenume);
                insertCommand.Parameters.AddWithValue("?", client.Telefon);
                insertCommand.Parameters.AddWithValue("?", client.CNP);
                insertCommand.Parameters.AddWithValue("?", client.Email);
                insertCommand.Parameters.AddWithValue("?", client.Judet);
                insertCommand.Parameters.AddWithValue("?", client.LocSec);
                insertCommand.Parameters.AddWithValue("?", client.Adresa);
                insertCommand.Parameters.AddWithValue("?", client.Venit);
                insertCommand.Parameters.AddWithValue("?", client.DataCreare);

                insertCommand.ExecuteNonQuery();
                MessageBox.Show("Client adăugat cu succes!");

                // Clear the text boxes after adding the client
                rtcNume.Text = "";
                rtcPrenume.Text = "";
                rtcTelefon.Text = "";
                rtcJudet.Text = "";
                rtcCNP.Text = "";
                rtcEmail.Text = "";
                rtcAdresa.Text = "";
                rtcLocSec.Text = "";
                conn.Close();
            }

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
