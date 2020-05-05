PPTXport
===

A PowerPoint Macro to export the content of a PPTX or PPTM for use with the [PptxGenJS](https://github.com/gitbrent/PptxGenJS) scipting library.

This is a very early first attempt but will continue to evolve - it's currently very rough, and will be optimized for specific use-cases that I need, but I would be more than happy to accept contributions to help make it more complete.

To use, create a new blank PPTX, import the module and run it. It will ask you to select the source file and then will process the elements and write out a new file called ppt.html.

If you have images in the PPTX or PPTM please create an [code]images[/code] directory before running the script.

Note: Currently on runs on Windows as it uses the Scripting.Dictionary object