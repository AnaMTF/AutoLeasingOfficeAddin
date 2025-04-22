using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

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
            // 1️⃣ – ce slide e activ?
            PowerPoint.Slide curSlide = null;

            // funcţionează în Normal view
            if (pptApp.ActiveWindow?.View?.Type ==
                PowerPoint.PpViewType.ppViewNormal)
            {
                curSlide = pptApp.ActiveWindow.View.Slide;
            }
            // fallback pentru SlideSorter view sau selecţie multiplă
            else if (pptApp.ActiveWindow?.Selection?.SlideRange?.Count > 0)
            {
                curSlide = pptApp.ActiveWindow.Selection.SlideRange[1];
            }

            if (curSlide == null)
            {
                MessageBox.Show("Selectaţi un slide cu ofertă.");
                return;
            }

            // 2️⃣ – citim tag‑ul cu OfferID
            string idText = curSlide.Tags["OfferID"];
            if (string.IsNullOrEmpty(idText))
            {
                MessageBox.Show("Slide‑ul selectat nu conţine ofertă asociată.");
                return;
            }

            int ofertaId = int.Parse(idText);

            // 3️⃣ – aici chemi logica ta (ex.: deschizi fereastră detalii,
            //     pui oferta într‑un contract Word, etc.)
            MessageBox.Show($"Ai selectat oferta cu ID = {ofertaId}");
        }
    }
}
