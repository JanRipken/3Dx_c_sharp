using System;
using System.Windows;

namespace MinimalCatia3DEXPERIENCE
{
    public class ExperienceControl
    {
        public ExperienceControl()
        {
            try
            {

                ExperienceConnection ec = new ExperienceConnection();

                // Finde Catia Prozess
                if (ec.CATIALaeuft())
                {
                    Console.WriteLine("0");

                    // Öffne ein neues Part
                    ec.Erzeuge3DShape();
                    Console.WriteLine("1");

                    // Erstelle eine Skizze
                    ec.ErstelleLeereSkizze();
                    Console.WriteLine("2");

                    // Generiere ein Profil
                    ec.ErzeugeKontur(40, 30);
                    Console.WriteLine("ZR 3");

                    // Extrudiere Balken
                    ec.ErzeugeZahnrad(5);
                    Console.WriteLine("ZR 4");
                }
                else
                {
                    Console.WriteLine("Laufende 3DEXPERIENCE-Applikation nicht gefunden");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception aufgetreten");
            }
            Console.WriteLine("Fertig - Taste drücken.");
            Console.ReadKey();
        }

        /// <summary>
        /// Dies ist der Einstiegspunkt des Programms.
        /// </summary>
        static void Main()
        {
            new ExperienceControl();
        }
    }
}
