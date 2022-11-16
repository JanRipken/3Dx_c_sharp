using CATPLMEnvBuild;
using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using PARTITF;
using ProductStructureClientIDL;
using System;
using System.Windows;

namespace MinimalCatia3DEXPERIENCE
{
    internal class ExperienceConnection
    {
        #region private Felder

        private INFITF.Application hsp_experienceApp;
        private Editor hsp_experienceEditor;
        private Sketch hsp_catiaProfil;

        private ShapeFactory SF;
        private HybridShapeFactory HSF;

        private Part myPart;
        private Sketches mySketches;

        #endregion

        #region master

        public bool CATIALaeuft()
        {
            try
            {
                object experienceObject = System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "CATIA.Application");
                hsp_experienceApp = (INFITF.Application)experienceObject;

                // Prüfen, ob auch wirklich 3DEXPERIENCE gefunden wurde.
                if (hsp_experienceApp.get_Name() != "3DEXPERIENCE")
                {
                    throw new Exception();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public void Erzeuge3DShape()
        {

            PLMNewService newService = (PLMNewService)hsp_experienceApp.GetSessionService("PLMNewService");

            // Ein "Editor"-Objekt wird als Outpout zurückgeliefert.
            newService.PLMCreate("3DShape", out hsp_experienceEditor);

            // Part-Objekt bekommen.
            myPart = (Part)hsp_experienceEditor.ActiveObject;

            // 3DShape bekommen.
            VPMRepReference myProductRepresentationReference = (VPMRepReference)myPart.Parent;

            #region Benennungen

            // ID
            myProductRepresentationReference.SetAttributeValue("PLM_ExternalID", "Hier eine ID eingeben");

            //Name
            // So geht es NICHT!
            // myProductRepresentationReference.set_Name("Hier einen Namen eingeben");

            // So geht es.
            myProductRepresentationReference.SetAttributeValue("Name", "Hier einen Namen eingeben");

            #endregion

            myPart.Update();
        }

        public void ErstelleLeereSkizze()
        {
            // Factories für das Erzeugen von Modellelementen (Std und Hybrid)
            SF = (ShapeFactory)myPart.ShapeFactory;
            HSF = (HybridShapeFactory)myPart.HybridShapeFactory;

            // geometrisches Set auswaehlen und umbenennen
            HybridBodies catHybridBodies1 = myPart.HybridBodies;
            HybridBody catHybridBody1;
            try
            {
                catHybridBody1 = catHybridBodies1.Item("Geometrisches Set.1");
            }
            catch (Exception)
            {
                MessageBox.Show("Kein geometrisches Set gefunden! " + Environment.NewLine +
                    "Ein PART manuell erzeugen und ein darauf achten, dass 'Geometisches Set' aktiviert ist.",
                    "Fehler", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            catHybridBody1.set_Name("Profile");

            // neue Skizze im ausgewaehlten geometrischen Set auf eine Offset Ebene legen
            mySketches = catHybridBody1.HybridSketches;
            OriginElements catOriginElements = myPart.OriginElements;
            Reference catReference1 = (Reference)catOriginElements.PlaneYZ;

            hsp_catiaProfil = mySketches.Add(catReference1);

            // Achsensystem in Skizze erstellen 
            ErzeugeAchsensystem();

            // Part aktualisieren
            myPart.Update();
        }

        private void ErzeugeAchsensystem()
        {
            object[] arr = new object[] {0.0, 0.0, 0.0,
                                         0.0, 1.0, 0.0,
                                         0.0, 0.0, 1.0 };
            hsp_catiaProfil.SetAbsoluteAxisData(arr);
        }

        #endregion

        #region Zahnrad

        public void ErzeugeKontur(int n, double radius)
        {
            // Skizze umbenennen
            hsp_catiaProfil.set_Name("ZR_Sketch");

            // Kreis in Skizze einzeichnen
            // Skizze oeffnen
            Factory2D catFactory2D1 = hsp_catiaProfil.OpenEdition();

            // Kreis erzeugen

            // erst die Punkte
            double alpha = 2 * Math.PI / n;
            double[] kreisPx = new double[n];
            double[] kreisPy = new double[n];
            Point2D[] catPoints = new Point2D[n];

            for (int ii = 0; ii < n; ii++)
            {
                kreisPx[ii] = Math.Cos(alpha * ii) * radius;
                kreisPy[ii] = Math.Sin(alpha * ii) * radius;

                catPoints[ii] = catFactory2D1.CreatePoint(kreisPx[ii], kreisPy[ii]);
            }

            // dann die Linien
            Line2D[] catLines = new Line2D[n];
            for (int ii = 1; ii < n; ii++) // Achtung, mit 1 starten wg. -1
            {
                catLines[ii] = catFactory2D1.CreateLine(kreisPx[ii - 1], kreisPy[ii - 1], kreisPx[ii], kreisPy[ii]);
                catLines[ii].StartPoint = catPoints[ii - 1];
                catLines[ii].EndPoint = catPoints[ii];
            }
            // Jetzt noch mit erste Linie [0] den Kreis schliessen
            catLines[0] = catFactory2D1.CreateLine(kreisPx[n - 1], kreisPy[n - 1], kreisPx[0], kreisPy[0]);
            catLines[0].StartPoint = catPoints[n - 1];
            catLines[0].EndPoint = catPoints[0];

            // Skizzierer verlassen
            hsp_catiaProfil.CloseEdition();
            // Part aktualisieren
            myPart.Update();
        }

        public void ErzeugeZahnrad(double dicke)
        {
            // Hauptkoerper in Bearbeitung definieren
            myPart.InWorkObject = myPart.MainBody;

            // Block(Balken) erzeugen
            //ShapeFactory catShapeFactory1 = (ShapeFactory)myPart.ShapeFactory;
            Pad catPad1 = SF.AddNewPad(hsp_catiaProfil, dicke);

            // Block umbenennen
            catPad1.set_Name("Zahnrad");

            // Part aktualisieren
            myPart.Update();
        }

        #endregion

    }
}


