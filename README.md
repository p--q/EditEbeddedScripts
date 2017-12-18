# EditEbeddedScripts

This is a PyDev project for editing embedded macros.

## System requirements

The confirmed environment is as follows.

LibreOffice 5.4 in Ubuntu 14.04 32bit

## USAGE

Place the only one Calc document you want to edit the embedded macro in the same hierarchy as the tools folder.

ex.  CalcDoc.ods

### Write embedded macro with replaceEmbeddedScripts.py

Put the macro you want to embed in the src folder in the same path as in the document.

ex.  src/Scripts/python

Execute replaceEmbeddedScripts.py in the tools folder.

This script writes Scripts/python folder to the Calc document and edits its manifest.xml.

### Extract embedded macro from document with getEmbeddedScripts.py

Execute getEmbeddedScripts.py in the tools folder.

Be careful as the contents of the src folder are replaced.
