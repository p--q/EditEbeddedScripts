# EditEbeddedScripts

This is a PyDev project for editing embedded macros.

## System requirements

The confirmed environment is as follows.

LibreOffice 5.4 in Ubuntu 14.04 32bit

## USAGE

Place the only one Calc document you want to edit the embedded macro in the same hierarchy as the tools folder (in other words,  <a href="https://github.com/p--q/EditEbeddedScripts/tree/master/EditEmbeddedScripts">EditEmbeddedScripts</a> folder).

ex.  CalcDoc.ods

There are many examples of embedded macros in <a href="https://sites.google.com/site/blogger2013pq/home/downloadfiles/calcexample">CalcExamples - p--q</a>.

### Write embedded macro with replaceEmbeddedScripts.py

Put the macro you want to embed in the src folder in the same path as in the document.

ex.  src/Scripts/python

Execute replaceEmbeddedScripts.py in the tools folder.

This script writes Scripts/python folder to the Calc document and edits its manifest.xml.

From version 0.1.3, close the document when it is open.

### Extract embedded macro from document with getEmbeddedScripts.py

Execute getEmbeddedScripts.py in the tools folder.

Be careful as the contents of the src folder are replaced.

### Release notes

2017-12-18 version 0.1.0 First release.

2017-12-21 version 0.1.1 Fixed a serious bug. Scripts/python folder in the document is registered in manifest.xml.

2017-12-21 version 0.1.2 Changed how to delete an existing the embedded macro folder.

2018-1-12 version 0.1.3 Commented out lines 40 - 42 of <a href="https://github.com/p--q/EditEbeddedScripts/blob/master/EditEmbeddedScripts/tools/replaceEmbeddedScripts.py#L40">replaceEmbeddedScripts.py</a>, since there are times when the embedded module is not updated unless LibreOffice's process is terminated.
