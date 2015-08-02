using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

class CSVRow : List<string>
{
}

class CSVFile : List<CSVRow>
{
   private Dictionary<string, int> clIndex;

   public void BuildIndex()
   {
      clIndex = new Dictionary<string, int>();

      if (Count == 0)
         throw new Exception("Kopfzeile nicht vorhanden");

      for (int i = 0; i < this[0].Count; ++i)
         clIndex.Add(this[0][i], i);
   }

   public string TryGetField(int row, string name)
   {
      try
      {
         return GetField(row, name);
      }
      catch (Exception)
      {
         return "";
      }
   }

   public string GetField(int row, string name)
   {
      if (clIndex == null)
         BuildIndex();

      if (row >= Count)
         throw new Exception("Zugriff auf nicht vorhandene Zeile: " + row);

      try
      {
         return this[row][clIndex[name]];
      }
      catch (Exception)
      {
         throw new Exception("Zugriff auf nicht vorhandene Spalte: " + name);
      }
   }

   public void SetField(int row, string name, string value)
   {
      GetField(row, name);

      this[row][clIndex[name]] = value;
   }
}

class CSVStreamReader : StreamReader
{
   private int clDelimiter;
   private int clEnclosure;

   public CSVStreamReader(string file, Encoding enc, char delimiter, char enclosure) : base(file, enc)
   {
      clDelimiter = (int) delimiter;
      clEnclosure = (int) enclosure;
   }

   public CSVRow ReadRow()
   {
      CSVRow row = new CSVRow();

      int c;
      string s = "";
      bool inEnclosure = false;

      while ((c = Read()) != -1)
      {
         if (inEnclosure)
         {
            if (c == clEnclosure)
            {
               inEnclosure = false;
            }
            else
            {
               s += (char) c;
            }
         }
         else
         {
            if (c == clEnclosure)
            {
               inEnclosure = true;
            }
            else if (c == clDelimiter)
            {
               row.Add(s);
               s = "";
            }
            else if (c == (int) '\n')
            {
               break;
            }
         }
      }

      if (inEnclosure)
         throw new Exception("Unerwartetes Datei-Ende");

      row.Add(s);
      return row;
   }
}

static class CSVFileReader
{
   public static CSVFile Open(string file, Encoding enc, char delimiter, char enclosure)
   {
      CSVFile csvObj = new CSVFile();

      CSVStreamReader csvsr = new CSVStreamReader(file, enc, delimiter, enclosure);

      while (! csvsr.EndOfStream)
         csvObj.Add(csvsr.ReadRow());

      return csvObj;
   }
}

class CSVFileWriter : StreamWriter
{
   private char clDelimiter;
   private char clEnclosure;

   public CSVFileWriter(string file, Encoding enc, char delimiter, char enclosure) : base(file, false, enc)
   {
      clDelimiter = delimiter;
      clEnclosure = enclosure;
   }

   private void WriteField(string field)
   {
      if (field == string.Empty)
         return;

      Write(clEnclosure);

      foreach (char c in field)
      {
         if (c == clEnclosure)
            Write('\\');

         Write(c);
      }

      Write(clEnclosure);
   }

   private void WriteRow(CSVRow row)
   {
      if (row.Count == 0)
         return;

      WriteField(row[0]);

      for (int i = 1; i < row.Count; ++i)
      {
         Write(clDelimiter);
         WriteField(row[i]);
      }
   }

   public void WriteCsv(CSVFile csvObj)
   {
      if (csvObj.Count == 0)
         return;

      WriteRow(csvObj[0]);

      for (int i = 1; i < csvObj.Count; ++i)
      {
         Write('\n');
         WriteRow(csvObj[i]);
      }

      Flush();
   }
}

class CSVFileConverter
{
   /*
      { "<Ausgabespalte>", "<Eingabespalte>" }

      Wenn Eingabespaltenname leer ("") ist, wird die Spalte leer gelassen oder mit einer
      speziellen Funktion befuellt.
   */

   private static Dictionary<string, string> clConfig = new Dictionary<string, string>()
   {
      { "Anrede", "Anrede" },
      { "Vorname", "Vorname" },
      { "Nachname", "Nachname" },
      { "Displayname", "" }, // = Vorname+Nachname

      { "Firma", "Firma" },
      { "Abteilung", "Abteilung" },
      { "Position", "Position" },

      { "Geburtstag", "Geburtstag" },

      { "PhotoId", "" },

      { "Kurzwahlnummer", ""},

      { "1. Typ Telefonnummer", "" },
      { "1. Telefonnummer", "" },
      { "2. Typ Telefonnummer", "" },
      { "2. Telefonnummer", "" },
      { "3. Typ Telefonnummer", "" },
      { "3. Telefonnummer", "" },
      { "4. Typ Telefonnummer", "" },
      { "4. Telefonnummer", "" },

      { "1. Typ von Emailadresse", "" },
      { "1. Emailadresse", "" },
      { "2. Typ von Emailadresse", "" },
      { "2. Emailadresse", "" },
      { "3. Typ von Emailadresse", "" },
      { "3. Emailadresse", "" },
      
      { "1. Typ Internetadresse", "" },
      { "1. Internetadresse", "" },
      { "2. Typ Internetadresse", "" },
      { "2. Internetadresse", "" },
      { "3. Typ Internetadresse", "" },
      { "3. Internetadresse", "" },
     
      { "1. Typ Adresse", "" },
      { "1. Strasse", "" },
      { "1. Postleitzahl", "" },
      { "1. Ort", "" },
      { "1. Staat", "" },
    
      { "2. Typ Adresse", "" },
      { "2. Strasse", "" },
      { "2. Postleitzahl", "" },
      { "2. Ort", "" },
      { "2. Staat", "" },
   
      { "3. Typ Adresse", "" },
      { "3. Strasse", "" },
      { "3. Postleitzahl", "" },
      { "3. Ort", "" },
      { "3. Staat", "" },

      { "", "" } // Gott, warum auch immer...
   };

   /*
      REGEX_EINGABESPALTE => AUSGABESPALTE
   */
   private static Dictionary<Regex, string> clPhoneNumberMatching = new Dictionary<Regex, string>() {
      { new Regex(@"haupttelefon", RegexOptions.IgnoreCase), "Haupttelefon" },
      { new Regex(@"telefon.*privat.*", RegexOptions.IgnoreCase), "Privat" },
      { new Regex(@"telefon.*(gesch.ftlich|firma)", RegexOptions.IgnoreCase), "Geschäftlich" },
      { new Regex(@"mobiltelefon", RegexOptions.IgnoreCase), "Mobil" }
   };

   public static CSVFile Convert(CSVFile inCsv)
   {
      CSVFile outCsv = new CSVFile {new CSVRow()};

      // Insert headline of outfile
      foreach (KeyValuePair<string, string> kvp in clConfig)
      {
         outCsv[0].Add(kvp.Key);
      }

      // Dictionary<Spaltenindex Eingabedatei, Spaltenname Ausgabedatei>
      Dictionary<int, String> clPhoneColumnMap = new Dictionary<int, String>();
      for (int i = 0; i < inCsv[0].Count; ++i)
      {
         foreach (KeyValuePair<Regex, string> kvp in clPhoneNumberMatching)
         {
            if (kvp.Key.IsMatch(inCsv[0][i]))
            {
               clPhoneColumnMap.Add(i, kvp.Value);
            }
         }
      }

      for (int r = 1; r < inCsv.Count; ++r)
      {
         outCsv.Add(new CSVRow());

         foreach (KeyValuePair<string, string> kvp in clConfig)
         {
            // leave column empty || fill it with function
            if (kvp.Value == String.Empty)
            {
               if (kvp.Key == "Displayname")
               {
                  outCsv[r].Add(inCsv.GetField(r, "Vorname") + " " + inCsv.GetField(r, "Nachname"));
               }
               else
                  outCsv[r].Add("");
            }
            // Simple copy of column
            else
            {
               outCsv[r].Add(inCsv.GetField(r, kvp.Value));
            }
         }

         // Phone-Number-Import
         int iPhoneNumber = 1;
         int maxPhoneNumber = 4;
         string phoneNumber;

         foreach (KeyValuePair<int, string> kvp in clPhoneColumnMap)
         {
            if ((phoneNumber = inCsv[r][kvp.Key]) == String.Empty)
               continue;

            outCsv.SetField(r, "" + iPhoneNumber + ". Typ Telefonnummer", kvp.Value);
            outCsv.SetField(r, "" + iPhoneNumber + ". Telefonnummer", Regex.Replace(phoneNumber, "\\D", ""));

            if (++iPhoneNumber > maxPhoneNumber)
               break;
         }
      }

      return outCsv;
   }
}

class Converter
{
   static void Main(string[] args)
   {
      string logFile = "log.txt";

      // delimiter config
      char inFileDelimiter = ',';
      char inFileEnclosure = '"';
      char outFileDelimiter = ';';
      char outFileEnclosure = '"';

      using (StreamWriter logWriter = File.AppendText(logFile))
      {
         try
         {
            if (args.Length == 0)
               throw new Exception("Das Programm benoetigt mindestens eine Eingabe-Datei!");

            foreach (string inFile in args)
            {
               CSVFile inCsv = CSVFileReader.Open(inFile,
                     Encoding.GetEncoding("iso-8859-1"),
                     inFileDelimiter,
                     inFileEnclosure
               );

               CSVFile outCsv = CSVFileConverter.Convert(inCsv);

               string outFile = inFile + ".konvertiert.csv";

               CSVFileWriter writer = new CSVFileWriter(outFile,
                     Encoding.GetEncoding("iso-8859-1"),
                     outFileDelimiter,
                     outFileEnclosure
               );

               writer.WriteCsv(outCsv);
            }
         }
         catch (Exception e)
         {
            logWriter.WriteLine(e.ToString());

            Console.WriteLine("Ein Fehler ist aufgetreten. Weitere Informationen befinden sich in der Log-Datei.\n\n> {0}\n", e.Message);
            Console.WriteLine("Enter zum Beenden druecken...");
            Console.ReadLine();
         }
      }
   }
}
