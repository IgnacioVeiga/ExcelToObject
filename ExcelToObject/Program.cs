using ExcelMapper;
using System;
using System.IO;
using System.Linq;

namespace ExcelToObject
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Iniciando programa...");
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); // obtiene la ruta del escritorio del usuario en sesion
            Console.Write("Indique el nombre de archivo (Formato .xlsx) ubicado en su escritorio: ");
            string fileName = Console.ReadLine();
            if (fileName.Remove(0, fileName.Length - 4) == ".xlsx") // si es formato ".xlsx" entra
            {
                using (var stream = File.OpenRead(path + "/" + fileName))
                using (var importer = new ExcelImporter(stream))
                {
                    ExcelSheet sheet = importer.ReadSheet();
                    Event[] events = sheet.ReadRows<Event>().ToArray();
                    Console.WriteLine(events[0].Name); // Pub Quiz
                    Console.WriteLine(events[1].Name); // Live Music
                    Console.WriteLine(events[2].Name); // Live Football
                }
            }
            else
            {
                Console.WriteLine("Solo archivos de formato .xlsx");
            }
            Console.ReadKey();
        }

        public enum EventCause
        {
            Profit,
            Charity
        }

        public class Event
        {
            public string Name { get; set; }
            public string Location { get; set; }
            public int Attendance { get; set; }
            public DateTime Date { get; set; }
            public Uri Link { get; set; }
            public EventCause Cause { get; set; }
        }
    }
}
