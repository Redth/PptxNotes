using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;

namespace PptxNotes
{
    public static class OpenXmlTools
    {
        public static List<string> ExportNotes (string pptFile)
        {
            var notes = new List<string> ();

            using (var ppt = PresentationDocument.Open (pptFile, false)) {
                var part = ppt.PresentationPart;

                var presentation = part.Presentation;

                foreach (var slideId in presentation.SlideIdList.Elements<SlideId> ()) {

                    var relId = slideId.RelationshipId;

                    // Get the slide part from the relationship ID.
                    var slide = (SlidePart)part.GetPartById (relId);

                    // Build a StringBuilder object.
                    var paragraphText = new System.Text.StringBuilder ();

                    if (slide.NotesSlidePart != null && slide.NotesSlidePart.NotesSlide != null)
                    {
                        foreach (var p in slide.NotesSlidePart.NotesSlide.Descendants<A.Paragraph> ())
                        {
                            foreach (var d in p.Descendants())
                            {                                
                                if (d is A.Run)
                                {
                                    var drun = d as A.Run;

                                    paragraphText.Append(drun.Text.Text);
                                }
                                else if (d is A.Break)
                                {
                                    paragraphText.AppendLine();
                                }
                            }

                            paragraphText.AppendLine();
                        }
                    }

                    notes.Add (paragraphText.ToString ());
                }
            }

            return notes;
        }

        public static void ImportNotes (string pptFile, List<string> notes)
        {
            using (var ppt = PresentationDocument.Open (pptFile, true)) {
                var part = ppt.PresentationPart;

                var presentation = part.Presentation;
                int slideCount = 0;

                foreach (var slideId in presentation.SlideIdList.Elements<SlideId> ()) {

                    if (notes.Count <= slideCount)
                        continue;

                    var note = notes [slideCount];
                    slideCount++;

                    var relId = slideId.RelationshipId;

                    // Get the slide part from the relationship ID.
                    var slide = (SlidePart)part.GetPartById (relId);

                    // See if we have a note already, if not create it
                    if (slide.NotesSlidePart == null) {
                        slide.AddNewPart<NotesSlidePart> (relId);
                    }
                    
                    var textBodyElems = new List<OpenXmlElement>();
                    textBodyElems.Add(new A.BodyProperties());
                    textBodyElems.Add(new A.ListStyle());

                    var strParagraphs = Regex.Split(note.Trim (), "\r?\n", RegexOptions.Multiline);

                    foreach (var strP in strParagraphs)
                    {
                        textBodyElems.Add(
                            new A.Paragraph(
                                new A.Run(
                                    new A.RunProperties { Language = "en-US", Dirty = false },
                                    new A.Text { Text = strP },
                                    new A.EndParagraphRunProperties { Language = "en-US", Dirty = false })));
                    }

                    var textBody = new TextBody(textBodyElems);
                    
                    var notesSlide = new NotesSlide (
                        new CommonSlideData (new ShapeTree (
                            new NonVisualGroupShapeProperties (
                                new NonVisualDrawingProperties () { Id = (UInt32Value)1U, Name = "" },
                                new NonVisualGroupShapeDrawingProperties (),
                                new ApplicationNonVisualDrawingProperties ()),
                            new GroupShapeProperties (new A.TransformGroup ()),
                            new Shape (
                                new NonVisualShapeProperties (
                                    new NonVisualDrawingProperties () { Id = (UInt32Value)2U, Name = "Slide Image Placeholder 1" },
                                    new NonVisualShapeDrawingProperties (new A.ShapeLocks () { NoGrouping = true, NoRotation = true, NoChangeAspect = true }),
                                    new ApplicationNonVisualDrawingProperties (new PlaceholderShape () { Type = PlaceholderValues.SlideImage })),
                                new ShapeProperties ()),
                            new Shape (
                            new NonVisualShapeProperties (
                                new NonVisualDrawingProperties () { Id = (UInt32Value)3U, Name = "Notes Placeholder 2" },
                                new NonVisualShapeDrawingProperties (new A.ShapeLocks () { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties (new PlaceholderShape () { Type = PlaceholderValues.Body, Index = (UInt32Value)1U })),
                            new ShapeProperties (),
                            textBody) // Use our TextBody from above that we constructed
                            )),
                        new ColorMapOverride (new A.MasterColorMapping ()));

                    slide.NotesSlidePart.NotesSlide = notesSlide;
                }

                ppt.Save ();
            }
        }
    }
}
