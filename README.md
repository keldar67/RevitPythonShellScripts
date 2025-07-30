# RevitPythonShellScripts

A collection of Random Revit Python Shell Scripts that do various things when run from the Revit Python Shell Tool.
These are mostly custom utilities or proofs of concept for more serious development.

## MaterialsToExcel.py
This script just iterates over all of the materials in the active model and collates all of the data that is usually found in the Material Editor on the Identity Tab and also the Graphics Tab and writes the data into Excel.
RGB Colours for Pattern Hatches are writtne as an RGB string as well as individual values, and the RGB Strings background colour is set to represent the colour specified by the values
