using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ToDoNotificheBSC
{
    public class Allegato
    {


        public Allegato(string filename, string descrizione, string path)
        {
            FileName = filename;
            Descrizione = descrizione;
            Path = path;
        }

        public string FileName { get; set; }
        public string Descrizione { get; set; }
        public string Path { get; set; }


    }

}
