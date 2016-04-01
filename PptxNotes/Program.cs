using System;
using System.Linq;
using System.Collections.Generic;

namespace PptxNotes
{
    class MainClass
    {
        static void Help ()
        {
            Console.WriteLine ("pptxnotes usage:");
            Console.WriteLine ("===================");

            Console.WriteLine ("Importing:");
            Console.WriteLine ("----------");
            Console.WriteLine ("pptnotes import ./notes.md ./presentation.pptx");
            Console.WriteLine ();

            Console.WriteLine ("Exporting:");
            Console.WriteLine ("----------");
            Console.WriteLine ("pptnotes export ./notes.md ./presentation.pptx");
            Console.WriteLine ();
        }

        public static void Main (string [] args)
        {
            if (args.Length < 3) {
                Console.WriteLine ("Invalid number of arguments");
                Help ();
                return;
            }

            var cmd = args [0];
            var notesFile = args [1];
            var pptxFile = args [2];

            if (cmd != "import" && cmd != "export") {
                Console.WriteLine ("Invalid Command, must use 'import' or 'export'");
                Help ();
                return;
            }

            if (cmd == "import" && !System.IO.File.Exists (notesFile)) {
                Console.WriteLine ("Cannot find notes file: " + notesFile);
                Help ();
                return;
            }

            if (!System.IO.File.Exists (pptxFile)) {
                Console.WriteLine ("Cannot find pptx file: " + pptxFile);
                Help ();
                return;
            }

            //var p = new Presentation (pptxFile);

            if (cmd == "import") {

                var notes = ReadNotes (notesFile);

                OpenXmlTools.ImportNotes (pptxFile, notes);
                
            } else if (cmd == "export") {

                var notes = OpenXmlTools.ExportNotes (pptxFile);

                WriteNotes (notes, notesFile);
            }
        }

        static void WriteNotes (List<string> notes, string notesFile)
        {
            int slide = 1;
            using (var f = System.IO.File.CreateText (notesFile)) {
                foreach (var n in notes) {
                    f.WriteLine ("### Slide " + slide);
                    f.WriteLine (n);
                    f.WriteLine ();
                    slide++;
                }
                f.Flush ();
            }
        }

        static List<string> ReadNotes (string notesFile)
        {
            var lines = System.IO.File.ReadAllLines (notesFile);
            var notes = new List<string> ();

            var noteOn = 0;
            var noteContent = string.Empty;

            foreach (var line in lines) {
                if (line.StartsWith ("###")) {
                    noteOn++;
                    if (noteOn > 1) {
                        notes.Add (noteContent);
                    }
                    noteContent = string.Empty;
                } else {
                    noteContent += line + Environment.NewLine;
                }
            }

            if (!string.IsNullOrEmpty (noteContent))
                notes.Add (noteContent);

            return notes;
        }        
    }
}
