# PptxNotes

Easily export notes from a .pptx to .md and Import into your .pptx from .md

NOTE: Unfortunately this only runs on windows as it uses Microsoft's OpenXml library, which kind of works on mono, but not for the bits needed in this app :(


## Exporting Notes

You can export notes from a .pptx file to a .md file:

```
pptxnotes.exe export notes.md presentation.pptx
```

The format they are exported in will be the same format used to import.  The exporter will add `Slide #` after each `###` separator, just to make things pretty.



## Importing Notes

You can import notes into a .pptx file from a given .md file:

```
pptxnotes.exe import notes.md presentation.pptx
```

Your `notes.md` should have lines starting with `###` to separate notes from each slide.  Anything after the '###' on the same line is ignored.  Here is an example:

```markdown
### Slide 1
These are the notes for your first slide.

### The second slide
Here are more notes

You can have multiple lines

### Etc...
Unfortunately, formatting is not preserved yet.

So only plain text work
 - Later I'll try and add formatting preservation
 - And make it more awesome...
```

