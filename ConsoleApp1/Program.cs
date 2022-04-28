using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

// Thanks to: https://youtu.be/9mUuJIKq40M

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Scrivo il percorso della cartella in cui sono contenuti i files dei fogli mensa digitali

            // il carattere @ serve per poter inserire nelle virgolette i caratteri '\' senza il carattere escape '\' quindi senza scrivere \\
            //String rootPath = @"\serverssn\Condivisa\MENSA_DIGITALE/";		
            //String rootPath = @"C:\Users\LENOVO M700\PROGRAMMAZIONE IN WINDOWS\C #\PROGETTI PERSONALI IN C#\CALCOLA TOTALE ALUNNI\CartellaDiProva\";
            String rootPath = System.IO.File.ReadAllText(@"pathToFolder.txt");

            var files = Directory.GetFiles(rootPath, "*.*", SearchOption.TopDirectoryOnly);
            Console.Beep();

            // Per i colori nel terminale, Thanks to: https://stackoverflow.com/questions/2743260/is-it-possible-to-write-to-the-console-in-colour-in-net
            Console.ForegroundColor = ConsoleColor.Yellow;

            // Scrivo la data
            // Thanks to: https://youtu.be/KKzSQ6r93dY
            // Thanks to: https://docs.microsoft.com/it-it/dotnet/standard/base-types/how-to-extract-the-day-of-the-week-from-a-specific-date
            //DateTime data = new DateTime();
            //var data = new DateTime();
            var data = DateTime.Now;  // Now è importantissimo! Altrimenti non otteniamo la data attuale ma dobbiamo impostarla noi
            List<String> giorniSettimana = new List<String> { "Lunedì", "Martedì", "Mercoledì", "Giovedì", "Sabato", "Domenica" };
            string giorno = data.DayOfWeek.ToString("d");
            string giornoDelMese = data.Day.ToString();
            string Mese = data.Month.ToString();
            var anno = data.Year;

            Console.WriteLine("Buongiorno, cara Valeria e Caro Francesco!");
            Console.WriteLine($"Oggi è {giorniSettimana[int.Parse(giorno) - 1]}, {giornoDelMese}/{Mese}/{anno}.");

            Console.WriteLine("\nPer prima cosa trasformiamo gli eventuali files.xls in files.xlsx, " +
                "\nper eliminare il messaggio di errore di Excel.");

            Console.WriteLine("\n\nPremere il tasto Invio per Continuare...");
            Console.ReadLine();
            Console.Beep();
            Console.Clear();

            // Trasformo i file.xls in xlsx
            foreach (string file in files)
            {
                string nomeFile = file.ToString();
                if (nomeFile.EndsWith("xls"))
                {
                    ChangeExtension(file);
                }
            }


            // aggiorno la variabile files in modo che contenga gli eventuali files trasformati in files.xlsx
            files = Directory.GetFiles(rootPath, "*.*", SearchOption.TopDirectoryOnly);

            // Creo una lista che contenga i nomi dei files.xlsx
            List<string> listaFiles = new List<String>();

            foreach (string file in files)
            {
                // Aggiungo alla listaFiles i nomi dei files.xlsx
                if (file.EndsWith("xlsx"))
                {
                    // Aggiungo i nomi dei files senza estensione alla listaFiles
                    // Questa lista mi servirà per calcolare il totale degli alunni che si fermano a mensa
                    listaFiles.Add(Path.GetFileNameWithoutExtension(file).ToString());
                }
            }

            // Creo un lista che contenga i nomi delle classi
            List<string> listaClassi = new List<string>();
            foreach (string classe in listaFiles)
            {
                listaClassi.Add(classe.Substring(37, 9));
            }

            // Stampo la lista delle classi per un controllo
            //foreach (string classe in listaClassi)
            //{
            //    Console.WriteLine(classe);
            //}

            Console.WriteLine(ControllaFogli(listaFiles));
            
            Console.WriteLine("Ecco i files presenti nella cartella MENSA_DIGITALE dopo la conversione del formato:\n\n");
            foreach (string file in files)
            {
                if (file.EndsWith("xlsx"))
                Console.WriteLine(Path.GetFileName(file));
            }

            Console.WriteLine("\n\nPremere il tasto Invio per Continuare...");
            Console.ReadLine();
            Console.Beep();
            Console.Clear();

            // Constrollo se tutti i fogli mensa sono stati salvati in MENSA_DIGITALE o se ce ne sono di meno o di più.
            // Questa variabile serve per capire se procedere con il calcolo del totale oppure no (in caso di un numero sbagliato di files mensa)
            Boolean isCheckOk = false;
            Console.WriteLine("Adesso controllo il numero di fogli mensa salvati in MENSA_DIGITALE");
            // Controllo se sono presenti i fogli mensa digitali di tutte le classi

            Console.WriteLine("\n\nPremere il tasto Invio per Continuare...");
            Console.ReadLine();
            Console.Beep();
            Console.Clear();

            if (ControllaFogli(listaFiles))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Nella cartella MENSA_DIGITALE sono presenti {listaFiles.Count} files");
                Console.WriteLine("Tutte le classi hanno salvato il foglio mensa digitale!");

                isCheckOk = true;
            }
            else if (listaFiles.Count < 13)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n\nNella cartella MENSA_DIGITALE sono presenti {listaFiles.Count} files");
                Console.WriteLine("\nAttenzione, non tutte le classi hanno compilato il foglio mensa digitale!");
                Console.WriteLine("\n\nNon posso procedere nel calcolo del totale degli alunni che si fermano a mensa.");
                isCheckOk = false;
            }
            else if (listaFiles.Count > 13)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n\nNella cartella MENSA_DIGITALE sono presenti {listaFiles.Count} files.");
                Console.WriteLine($"Attenzione, in qualche classe hanno salvato il foglio mensa digitale più volte!");
                Console.WriteLine("\nNon posso procedere nel calcolo del totale degli alunni che si fermano a mensa.");
                isCheckOk = false;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Nella cartella MENSA_DIGITALE sono presenti {listaFiles.Count} files.");
                Console.WriteLine($"\nE' possibile, però, che una classe abbia salvato più volte il foglio mensa, " +
                    $"\n\nmentre qualche altra classe non lo abbia salvato.");
                Console.WriteLine("\nNon posso procedere con il calcolo del totale degli alunni che si fermano a mensa.");
                isCheckOk = false;
            }

            // Calcolo il totale degli alunni che si fermano a pranzo
            if (isCheckOk)
            {
                // Riscrivo la data
                Console.WriteLine($"\nOggi è {giorniSettimana[int.Parse(giorno) - 1]}, {giornoDelMese}/{Mese}/{anno}.");
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                
                Console.WriteLine("\nEcco il totale degli alunni che si fermano a pranzo oggi in base ai files salvati:\n");
                List<int> listaNumeri = new List<int>();
                foreach (string file in listaFiles)
                {
                    string numeroAlunni;
                    if (file.Length == 58)
                    {
                        numeroAlunni = file.Substring(file.Length - 2);
                        listaNumeri.Add(int.Parse(numeroAlunni));
                    }
                    else
                    {
                        numeroAlunni = file.Substring(file.Length - 1);
                        listaNumeri.Add(int.Parse(numeroAlunni));
                    }
                }

                // Stampo l'elenco delle classi e il totale degli alunni che si fermano a mensa per ogni classe
                for (int i = 0; i < 13; i++)
                {
                    Console.WriteLine($"{listaFiles[i].Substring(37, 9)}{listaNumeri[i], 20}"); // il numero dopo la virgola allinea a destra dopo 49 spazi

                }
                //Calcolo il totale degli alunni
                int totale = 0;
                foreach (int numero in listaNumeri)
                {
                    totale += numero;
                }


                Console.WriteLine("                          ---");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"TOTALE: {totale, 21}");  // il numero dopo la virgola allinea a destra dopo 50 spazi
                Console.ForegroundColor = ConsoleColor.DarkCyan;

                // Calcolo i vari totali degli alunni a seconda del giorno della settimana
                int giornoDaControllare = int.Parse(giorno);
                giornoDaControllare = 1;

                switch (giornoDaControllare)
                {
                    case 1:
                        Console.WriteLine("\n\nMENSA PICCOLA - primo turno");
                        Console.WriteLine($"{listaClassi[0]}{listaNumeri[0], 20}");
                        Console.WriteLine($"{listaClassi[1]}{listaNumeri[1], 20}");
                        Console.WriteLine($"{listaClassi[2]}{listaNumeri[2], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[0] + listaNumeri[1] + listaNumeri[2], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nMENSA PICCOLA - secondo turno");
                        Console.WriteLine($"{listaClassi[4]}{listaNumeri[4], 20}");
                        Console.WriteLine($"{listaClassi[5]}{listaNumeri[5], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[4] + listaNumeri[5], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - primo turno");
                        Console.WriteLine($"{listaClassi[3]}{listaNumeri[3], 20}");
                        Console.WriteLine($"{listaClassi[6]}{listaNumeri[6], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[3] + listaNumeri[6], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - secondo turno");
                        Console.WriteLine($"{listaClassi[7]}{listaNumeri[7], 20}");
                        Console.WriteLine($"{listaClassi[8]}{listaNumeri[8], 20}");
                        Console.WriteLine($"{listaClassi[9]}{listaNumeri[9], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[7] + listaNumeri[8] + listaNumeri[9], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nNELLE CLASSI");
                        Console.WriteLine($"{listaClassi[10]}{listaNumeri[10], 20}");
                        Console.WriteLine($"{listaClassi[11]}{listaNumeri[11], 20}");
                        Console.WriteLine($"{listaClassi[12]}{listaNumeri[12], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[10] + listaNumeri[11] + listaNumeri[12], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        break;
                    case 2:
                        Console.WriteLine("\n\nMENSA PICCOLA - primo turno");
                        Console.WriteLine($"{listaClassi[0]}{listaNumeri[0], 20}");
                        Console.WriteLine($"{listaClassi[1]}{listaNumeri[1], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[0] + listaNumeri[1], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nMENSA PICCOLA - secondo turno");
                        Console.WriteLine($"{listaClassi[8]}{listaNumeri[8], 20}");
                        Console.WriteLine($"{listaClassi[9]}{listaNumeri[9], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[8] + listaNumeri[9], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - primo turno");
                        Console.WriteLine($"{listaClassi[2]}{listaNumeri[2], 20}");
                        Console.WriteLine($"{listaClassi[3]}{listaNumeri[3], 20}");
                        Console.WriteLine($"{listaClassi[4]}{listaNumeri[4], 20}");
                        Console.WriteLine($"{listaClassi[5]}{listaNumeri[5], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[2] + listaNumeri[3] + listaNumeri[4] + listaNumeri[5], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - secondo turno");
                        Console.WriteLine($"{listaClassi[6]}{listaNumeri[6], 20}");
                        Console.WriteLine($"{listaClassi[7]}{listaNumeri[7], 20}");
                        Console.WriteLine($"{listaClassi[10]}{listaNumeri[10], 20}");
                        Console.WriteLine($"{listaClassi[11]}{listaNumeri[11], 20}");
                        Console.WriteLine($"{listaClassi[12]}{listaNumeri[12], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[6] + listaNumeri[7] + listaNumeri[10] + listaNumeri[11] + listaNumeri[12], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        break;
                    case 3:
                        Console.WriteLine("\n\nMENSA PICCOLA - primo turno");
                        Console.WriteLine($"{listaClassi[0]}{listaNumeri[0], 20}");
                        Console.WriteLine($"{listaClassi[1]}{listaNumeri[1], 20}");
                        Console.WriteLine($"{listaClassi[3]}{listaNumeri[3], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[0] + listaNumeri[1] + listaNumeri[3], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nMENSA PICCOLA - secondo turno");
                        Console.WriteLine($"{listaClassi[8]}{listaNumeri[9], 20}");
                        Console.WriteLine($"{listaClassi[9]}{listaNumeri[9], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[8] + listaNumeri[9], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - primo turno");
                        Console.WriteLine($"{listaClassi[5]}{listaNumeri[5], 20}");
                        Console.WriteLine($"{listaClassi[2]}{listaNumeri[2], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[5] + listaNumeri[2], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - secondo turno");
                        Console.WriteLine($"{listaClassi[4]}{listaNumeri[4], 20}");
                        Console.WriteLine($"{listaClassi[6]}{listaNumeri[6], 20}");
                        Console.WriteLine($"{listaClassi[7]}{listaNumeri[7], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[4] + listaNumeri[6] + listaNumeri[7], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nNELLE CLASSI");
                        Console.WriteLine($"{listaClassi[10]}{listaNumeri[10], 20}");
                        Console.WriteLine($"{listaClassi[11]}{listaNumeri[11], 20}");
                        Console.WriteLine($"{listaClassi[12]}{listaNumeri[12], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[10] + listaNumeri[11] + listaNumeri[12], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        break;
                    case 4:
                        Console.WriteLine("\n\nMENSA PICCOLA - primo turno");
                        Console.WriteLine($"{listaClassi[0]}{listaNumeri[0], 20}");
                        Console.WriteLine($"{listaClassi[1]}{listaNumeri[1], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[0] + listaNumeri[1], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nMENSA PICCOLA - secondo turno");
                        Console.WriteLine($"{listaClassi[6]}{listaNumeri[6], 20}");
                        Console.WriteLine($"{listaClassi[7]}{listaNumeri[7], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[6] + listaNumeri[7], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - primo turno");
                        Console.WriteLine($"{listaClassi[2]}{listaNumeri[2], 20}");
                        Console.WriteLine($"{listaClassi[3]}{listaNumeri[3], 20}");
                        Console.WriteLine($"{listaClassi[8]}{listaNumeri[8], 20}");
                        Console.WriteLine($"{listaClassi[9]}{listaNumeri[9], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[2] + listaNumeri[3] + listaNumeri[8] + listaNumeri[9], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - secondo turno");
                        Console.WriteLine($"{listaClassi[4]}{listaNumeri[4], 20}");
                        Console.WriteLine($"{listaClassi[5]}{listaNumeri[5], 20}");
                        Console.WriteLine($"{listaClassi[10]}{listaNumeri[10], 20}");
                        Console.WriteLine($"{listaClassi[11]}{listaNumeri[11], 20}");
                        Console.WriteLine($"{listaClassi[12]}{listaNumeri[12], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[4] + listaNumeri[5] + listaNumeri[10] + listaNumeri[11] + listaNumeri[12], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        break;
                    case 5:
                        Console.WriteLine("\n\nMENSA PICCOLA - primo turno");
                        Console.WriteLine($"{listaClassi[0]}{listaNumeri[0], 20}");
                        Console.WriteLine($"{listaClassi[1]}{listaNumeri[1], 20}");
                        Console.WriteLine($"{listaClassi[2]}{listaNumeri[2], 20}");
                        Console.WriteLine($"{listaClassi[3]}{listaNumeri[3], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[0] + listaNumeri[1] + listaNumeri[2] + listaNumeri[3], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nMENSA PICCOLA - secondo turno");
                        Console.WriteLine($"{listaClassi[6]}{listaNumeri[6], 20}");
                        Console.WriteLine($"{listaClassi[7]}{listaNumeri[7], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[6] + listaNumeri[7], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - primo turno");
                        Console.WriteLine($"{listaClassi[4]}{listaNumeri[4], 20}");
                        Console.WriteLine($"{listaClassi[5]}{listaNumeri[5], 20}");
                        Console.WriteLine($"{listaClassi[10]}{listaNumeri[10], 20}");
                        Console.WriteLine($"{listaClassi[11]}{listaNumeri[11], 20}");
                        Console.WriteLine($"{listaClassi[12]}{listaNumeri[12], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[5] + listaNumeri[2] + listaNumeri[8] + listaNumeri[9] + listaNumeri[10] + listaNumeri[11] + listaNumeri[12], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;

                        Console.WriteLine("\n\nSALONE - secondo turno");
                        Console.WriteLine($"{listaClassi[8]}{listaNumeri[8], 20}");
                        Console.WriteLine($"{listaClassi[9]}{listaNumeri[9], 20}");
                        Console.WriteLine($"                          ---");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"TOTALE: {listaNumeri[8] + listaNumeri[9], 21}");
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        break;

                }

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\n\nBuona Giornata!");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nRiprovare in un secondo momento quando i files saranno stati salvati correttamente.");
            }

            // Per evitare che si chiuda la finestra del terminale
            Console.WriteLine("\n\nPremere il tasto Invio per chiudere il programma");
            Console.ReadLine();
        }


        // Thanks to: https://youtu.be/5UFs2-X4tjc?list=PLpwlV5BeFjIjMKvEuJxD5UiJ-2G2QzpYO
        // Questa funzione serve per cambiare l'estensione di un file in un'altra estensione
        // In questo caso un file.xls viene trasformato in un file.xlsx
        public static void ChangeExtension(string oldExtension)
        {
            string newExtension = oldExtension + "x";
            File.Move(oldExtension, newExtension);
        }

        // Funzione per controllare che tutte le classi abbiano salvato il foglio mensa
        public static Boolean ControllaFogli(List<String> lista)
        {
            Boolean risultatoControllo = false;
            // Se al termine del controllo il counter sarà arrivato a 13 allora vorrà
            // dire che tutte le classi hanno salvato il foglio
            int counter = 0;
            // Controllo che tutte le classi abbiano salvato il file
            if (lista.Count == 13)
            {

                if (lista[0].Substring(37, 9) == "Classe 1A")
                {
                    counter++;
                }
                if (lista[1].Substring(37, 9) == "Classe 1B")
                {
                    counter++;
                }
                if (lista[2].Substring(37, 9) == "Classe 2A")
                {
                    counter++;
                }
                if (lista[3].Substring(37, 9) == "Classe 2B")
                {
                    counter++;
                }
                if (lista[4].Substring(37, 9) == "Classe 3A")
                {
                    counter++;
                }
                if (lista[5].Substring(37, 9) == "Classe 3B")
                {
                    counter++;
                }
                if (lista[6].Substring(37, 9) == "Classe 4A")
                {
                    counter++;
                }
                if (lista[7].Substring(37, 9) == "Classe 4B")
                {
                    counter++;
                }
                if (lista[8].Substring(37, 9) == "Classe 5A")
                {
                    counter++;
                }
                if (lista[9].Substring(37, 9) == "Classe 5B")
                {
                    counter++;
                }
                if (lista[10].Substring(37, 9) == "Classe M1")
                {
                    counter++;
                }
                if (lista[11].Substring(37, 9) == "Classe M2")
                {
                    counter++;
                }
                if (lista[12].Substring(37, 9) == "Classe M3")
                {
                    counter++;
                }

                if (counter == 13)
                {
                    risultatoControllo = true;
                }
            }
            

            return risultatoControllo;
        }



    }
}
