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
using System.IO.Pipes;
using Microsoft.Office.Interop.PowerPoint;


namespace WordClientNou
{
    public partial class ThisDocument
    {
        Button btnAdaugareClient;
        Button btnPowerpoint;
        
        string connString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\IPLAproiect.accdb";
        string ImageFolder =@"C:\Users\anama\source\repos\ProiectIPLA\WordClientNou\ImaginiDB";

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

            // la deschiderea documentului, clientul actual este resetat
            var conn = new OleDbConnection(connString);
            conn.Open();
            const string sql = "UPDATE DateActuale SET IdClientActual = -1, IdOfertaActuala = -1 WHERE ID = 1";
            using (var cmd = new OleDbCommand(sql, conn))
            {
                cmd.ExecuteNonQuery();
            }
            conn.Close();
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
                AddOfertaSlide(pres, oferta);

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
                            NumePoza = r.GetString(r.GetOrdinal("Poza"))
                        };
                        list.Add(o);
                    }
                }
            }
            return list;
        }

        private void AddOfertaSlide(PowerPoint.Presentation pres, Oferta of)
        {
            var sld = pres.Slides.Add(pres.Slides.Count + 1,
                        PowerPoint.PpSlideLayout.ppLayoutBlank);
            sld.Tags.Add("OfferID", of.IdOferta.ToString());

            // 🔷 Gradient albastru pe fundal
            sld.FollowMasterBackground = Office.MsoTriState.msoFalse;
            var bgFill = sld.Background.Fill;
            bgFill.Visible = Office.MsoTriState.msoTrue;
            bgFill.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
            bgFill.BackColor.RGB = ColorTranslator.ToOle(Color.White);
            bgFill.TwoColorGradient(Office.MsoGradientStyle.msoGradientFromCenter, 1);
            bgFill.GradientStops.Insert(ColorTranslator.ToOle(Color.DeepSkyBlue), 0.0f);
            bgFill.GradientStops.Insert(ColorTranslator.ToOle(Color.White), 1.0f);

            // 🔷 Titlu
            var titluBox = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 30, 700, 60);
            var rngTitlu = titluBox.TextFrame.TextRange;
            rngTitlu.Text = $"{of.Marca} {of.NumeMasina}";
            rngTitlu.Font.Size = 42;
            rngTitlu.Font.Bold = Office.MsoTriState.msoTrue;
            rngTitlu.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            rngTitlu.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            titluBox.Fill.Visible = Office.MsoTriState.msoTrue;
            titluBox.Fill.Solid();
            titluBox.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(173, 216, 230)); // baby blue

            // 🔷 Imagine maşină - rounded rectangle cu imagine de fundal
            string imgPath = System.IO.Path.Combine(ImageFolder, of.NumePoza ?? "");
            if (File.Exists(imgPath))
            {
                var shape = sld.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                                                 160, 110, 600, 300);
                shape.Fill.UserPicture(imgPath);
                shape.Line.Visible = Office.MsoTriState.msoFalse;
            }

            // 🔷 Stea preț – colț dreapta
            var stea = sld.Shapes.AddShape(Office.MsoAutoShapeType.msoShape5pointStar,
                                           670, 110, 160, 160);
            stea.Fill.Solid();
            stea.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.DarkBlue);
            stea.Line.Visible = Office.MsoTriState.msoFalse;

            // 🔷 Caseta cu preț – mai lată, centrată
            var pretBox = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                680, 160, 140, 40); // ⬅️ lățit
            pretBox.TextFrame.TextRange.Text = of.PretZi.ToString("0");
            pretBox.TextFrame.TextRange.Font.Size = 32;
            pretBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            pretBox.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.HotPink);
            pretBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignCenter;
            pretBox.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle; // ⬅️ centrare verticală

            // 🔷 Caseta cu „Lei” – puțin mai jos, tot centrată
            var leiBox = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                685, 195, 130, 30); // ⬅️ lățit și repoziționat
            leiBox.TextFrame.TextRange.Text = "Lei";
            leiBox.TextFrame.TextRange.Font.Size = 24;
            leiBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            leiBox.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.HotPink);
            leiBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignCenter;
            leiBox.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

            // 🔷 Detalii sub imagine
            var detalii = sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 100, 430, 700, 60);
            detalii.Fill.Visible = Office.MsoTriState.msoTrue;
            detalii.Fill.Solid();
            detalii.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(208, 240, 255));

            var rngDet = detalii.TextFrame.TextRange;
            rngDet.Text = $"Durată minimă: {of.DurataMinima} zile | Tarif: {of.PretZi:C}/zi\n" +
                          $"Carburant: {of.TipCarburant} | Locuri: {of.NrLocuri} | Km incluși/zi: {of.Kilometraj}";
            rngDet.Font.Size = 18;
            rngDet.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);
            rngDet.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
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
                Client client = new Client(rtcNume.Text, rtcPrenume.Text, rtcTelefon.Text, rtcJudet.Text, rtcCNP.Text, rtcEmail.Text, rtcAdresa.Text, rtcLocSec.Text, dlcVenit.Text);
                try
                {
                    var conn = new OleDbConnection(connString);
                    conn.Open();

                    /* 1️⃣  INSERT fără coloana-cheie (Id este AutoNumber) */
                    const string sqlIns =
                        @"INSERT INTO Clienti
              (Nume, Prenume, Telefon, CNP, Email,
               Judet, LocSec, Adresa, Venit, DataCreare)
              VALUES (?,?,?,?,?,?,?,?,?,?)";

                    using (var cmdIns = new OleDbCommand(sqlIns, conn))
                    {
                        cmdIns.Parameters.AddWithValue("?", rtcNume.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcPrenume.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcTelefon.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcCNP.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcEmail.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcJudet.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcLocSec.Text);
                        cmdIns.Parameters.AddWithValue("?", rtcAdresa.Text);
                        cmdIns.Parameters.AddWithValue("?", dlcVenit.Text);
                        cmdIns.Parameters.AddWithValue("?", DateTime.Now);
                        cmdIns.ExecuteNonQuery();
                    }

                    /* 2️⃣  ID-ul nou – concurrent-safe cu @@IDENTITY  */
                    int newClientId;
                    using (var cmdId = new OleDbCommand("SELECT @@IDENTITY", conn))
                    {
                        newClientId = System.Convert.ToInt32(cmdId.ExecuteScalar());
                    }

                    /* 3️⃣  Scriem contextul comun  */
                    const string sqlCtx =
                        "UPDATE DateActuale SET IdClientActual = ? WHERE ID = 1";

                    using (var cmdCtx = new OleDbCommand(sqlCtx, conn))
                    {
                        cmdCtx.Parameters.AddWithValue("?", newClientId);
                        cmdCtx.ExecuteNonQuery();
                    }

                    MessageBox.Show("Client adăugat cu succes!");

                    /* 4️⃣  Golește controalele */
                    rtcNume.Text = rtcPrenume.Text = rtcTelefon.Text = rtcJudet.Text =
                    rtcCNP.Text = rtcEmail.Text = rtcAdresa.Text = rtcLocSec.Text = "";

                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Eroare la salvarea clientului: " + ex.Message);
                }

                
            }

        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            //pipe.Close();
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
