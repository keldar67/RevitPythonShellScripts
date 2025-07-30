import os
import datetime
from pathlib import Path

#Import the Excel bits and pieces.
import clr
clr.AddReference('Microsoft.Office.Interop.Excel')
from Microsoft.Office.Interop import Excel
clr.AddReference("System.Drawing")
from System.Drawing import Color, ColorTranslator
    
excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.Workbooks.Add()
sheet = workbook.Sheets[1]

# Set to True to write out to EXCEL
OUTPUTTOEXCEL = False

# Define EXCEL Columns for Output:
COL_ELEMENTID         = 1
COL_MATERIALNAME      = 2
COL_DESCRIPTION       = 3
COL_MATCLASS          = 4
COL_COMMENTS          = 5
COL_KEYWORDS          = 6
COL_MANUFACTURER      = 7
COL_MODEL             = 8
COL_COST              = 9
COL_URL               = 10
COL_KEYNOTE           = 11
COL_MARK              = 12
COL_USERENDAPP        = 13
COL_COLOR             = 14
COL_COLOR_R           = 15
COL_COLOR_G           = 16
COL_COLOR_B           = 17
COL_TRANSPARENCY      = 18
COL_SURFFGPATTERN     = 19
COL_SURFFGCOLOR       = 20
COL_SURFFGCOLOR_R     = 21
COL_SURFFGCOLOR_G     = 22
COL_SURFFGCOLOR_B     = 23
COL_SURFBGPATTERN     = 24
COL_SURFBGCOLOR       = 25
COL_SURFBGCOLOR_R     = 26
COL_SURFBGCOLOR_G     = 27
COL_SURFBGCOLOR_B     = 28
COL_CUTFGPATTERN      = 29
COL_CUTFGCOLOR        = 30
COL_CUTFGCOLOR_R      = 31
COL_CUTFGCOLOR_G      = 32
COL_CUTFGCOLOR_B      = 33
COL_CUTBGPATTERN      = 34
COL_CUTBGCOLOR        = 35
COL_CUTBGCOLOR_R      = 36
COL_CUTBGCOLOR_G      = 37
COL_CUTBGCOLOR_B      = 38


#-------------------------------------------------------------------------------------
def getPatternName(id):
  if id and not id == ElementId.InvalidElementId:
    print(id)
    return Element.Name.GetValue(doc.GetElement(id))
  else:
    return ""
#-------------------------------------------------------------------------------------
def getMaterialClass(m):
  if m.MaterialClass == None:
    return ""
  else:
    return m.MaterialClass
#-------------------------------------------------------------------------------------
def getParamAsString(elem, param):
  p = elem.LookupParameter(param)
  if p == None or False == p.HasValue:
    return ""
  else:
    return elem.LookupParameter(param).AsString()
#-------------------------------------------------------------------------------------
def getRGBText(r, g, b):
  red   = str(r).PadLeft(3,'0')
  green = str(g).PadLeft(3,'0')
  blue  = str(b).PadLeft(3,'0')
  
  return 'RGB ' + red + '-' + green + '-' + blue
  
#-------------------------------------------------------------------------------------
def WriteHeaders(sheet):
  row = 1
  sheet.Cells[row,COL_ELEMENTID].Value2        = 'Element Id'
  sheet.Cells[row,COL_MATERIALNAME].Value2     = 'Material Name'
  sheet.Cells[row,COL_DESCRIPTION].Value2      = 'Description'
  sheet.Cells[row,COL_MATCLASS].Value2         = 'Material Class'
  sheet.Cells[row,COL_COMMENTS].Value2         = 'Comments'
  sheet.Cells[row,COL_KEYWORDS].Value2         = 'Keywords'
  sheet.Cells[row,COL_MANUFACTURER].Value2     = 'Manufacturer'
  sheet.Cells[row,COL_MODEL].Value2            = 'Model'
  sheet.Cells[row,COL_COST].Value2             = 'Cost'
  sheet.Cells[row,COL_URL].Value2              = 'URL'
  sheet.Cells[row,COL_KEYNOTE].Value2          = 'Keynote'
  sheet.Cells[row,COL_MARK].Value2             = '# Mark'
  sheet.Cells[row,COL_USERENDAPP].Value2       = 'Use Render Appearance'
  sheet.Cells[row,COL_COLOR].Value2            = 'Material Color'
  sheet.Cells[row,COL_COLOR_R].Value2          = 'R'
  sheet.Cells[row,COL_COLOR_G].Value2          = 'G'
  sheet.Cells[row,COL_COLOR_B].Value2          = 'B'
  sheet.Cells[row,COL_TRANSPARENCY].Value2     = 'Transparency'
  sheet.Cells[row,COL_SURFFGPATTERN].Value2    = 'Surface FG Pattern'
  sheet.Cells[row,COL_SURFFGCOLOR].Value2      = 'Surface FG Colour'
  sheet.Cells[row,COL_SURFFGCOLOR_R].Value2    = 'R'
  sheet.Cells[row,COL_SURFFGCOLOR_G].Value2    = 'G'
  sheet.Cells[row,COL_SURFFGCOLOR_B].Value2    = 'B'
  sheet.Cells[row,COL_SURFBGPATTERN].Value2    = 'Surface BG Pattern'
  sheet.Cells[row,COL_SURFBGCOLOR].Value2      = 'Surface BG Color'
  sheet.Cells[row,COL_SURFBGCOLOR_R].Value2    = 'R'
  sheet.Cells[row,COL_SURFBGCOLOR_G].Value2    = 'G'
  sheet.Cells[row,COL_SURFBGCOLOR_B].Value2    = 'B'
  sheet.Cells[row,COL_CUTFGPATTERN].Value2     = 'Cut FG Pattern'
  sheet.Cells[row,COL_CUTFGCOLOR].Value2       = 'Cut FG Color'
  sheet.Cells[row,COL_CUTFGCOLOR_R].Value2     = 'R'
  sheet.Cells[row,COL_CUTFGCOLOR_G].Value2     = 'G'
  sheet.Cells[row,COL_CUTFGCOLOR_B].Value2     = 'B'
  sheet.Cells[row,COL_CUTBGPATTERN].Value2     = 'Cut BG Pattern'
  sheet.Cells[row,COL_CUTBGCOLOR].Value2       = 'Cut BG Color'
  sheet.Cells[row,COL_CUTBGCOLOR_R].Value2     = 'R'
  sheet.Cells[row,COL_CUTBGCOLOR_G].Value2     = 'G'
  sheet.Cells[row,COL_CUTBGCOLOR_B].Value2     = 'B'

  
  
#-------------------------------------------------------------------------------------
# Get the documents materials
theMaterials = FilteredElementCollector(doc).OfClass(Material).WhereElementIsNotElementType()

# Color Stuff
_black = Color.FromArgb(0,0,0)
_white = Color.FromArgb(255,255,255)

# Initialise the Excel Row Counter
row = 1

WriteHeaders(sheet)

# The Main Loop through the materials:
for aMaterial in theMaterials:
  
  # Move to the next Row
  row += 1
  
  print('+-----------------------------+')
  
  #-------------------------------------------------------------------------------------
  # Get the Identity Tab information
  #-------------------------------------------------------------------------------------
  print(aMaterial.Name + ' [' + str(aMaterial.Id) + ']')
  theDescription = getParamAsString(aMaterial, "Description")
  theClass = getMaterialClass(aMaterial)
  theComments = getParamAsString(aMaterial, 'Comments')
  theKeywords = getParamAsString(aMaterial, 'Keywords')
  
  sheet.Cells[row,COL_ELEMENTID].Value2    = str(aMaterial.Id)
  sheet.Cells[row,COL_MATERIALNAME].Value2 = aMaterial.Name
  sheet.Cells[row,COL_DESCRIPTION].Value2  = theDescription
  sheet.Cells[row,COL_MATCLASS].Value2     = theClass
  sheet.Cells[row,COL_COMMENTS].Value2     = theComments
  sheet.Cells[row,COL_KEYWORDS].Value2     = theKeywords
  
  print('Description: ' + theDescription)
  print('Class: ' + theClass)
  print('Comments: ' + theComments)
  print('Keywords: ' + theKeywords)
  #-------------------------------------------------------------------------------------
  # Get the Graphics Tab information
  #-------------------------------------------------------------------------------------
  print('----Graphics----')
  ura = aMaterial.UseRenderAppearanceForShading
  R = aMaterial.Color.Red
  G = aMaterial.Color.Green
  B = aMaterial.Color.Blue
  col = getRGBText(R, G, B)
  trans = aMaterial.Transparency
  cost = aMaterial.LookupParameter('Cost').AsDouble()
  url = getParamAsString(aMaterial, 'URL')
  mfr = getParamAsString(aMaterial, 'Manufacturer')
  mod = getParamAsString(aMaterial, 'Model')
  
  print('Material = [' + str(R) + ':' + str(G) + ':' + str(B) + '] ' + str(ura))
  
  sheet.Cells[row,COL_USERENDAPP].Value2    = ura
  sheet.Cells[row,COL_COLOR].Value2         = col
  sheet.Cells[row,COL_COLOR_R].Value2       = str(R)
  sheet.Cells[row,COL_COLOR_G].Value2       = str(G)
  sheet.Cells[row,COL_COLOR_B].Value2       = str(B)
  sheet.Cells[row,COL_TRANSPARENCY].Value2  = trans
  sheet.Cells[row,COL_COST].Value2          = cost
  sheet.Cells[row,COL_URL].Value2           = url
  sheet.Cells[row,COL_MANUFACTURER].Value2  = mfr
  sheet.Cells[row,COL_MODEL].Value2         = mod
  
  #Shade the background of the colour cell to match the RGB value
  sheet.Cells[row,COL_COLOR].Interior.Color = Color.FromArgb(R,G,B)
  #Determine how dark the colour is and set the font color accordingly
  textcol = _black if (R * 0.299 + G * 0.587 + B * 0.114) > 186 else _white
  sheet.Cells[row,COL_COLOR].Font.Color = textcol
  
  print('Cut Pattern')
  # Get the Background Pattern Info
  id = aMaterial.CutBackgroundPatternId
  R = aMaterial.CutBackgroundPatternColor.Red
  G = aMaterial.CutBackgroundPatternColor.Green
  B = aMaterial.CutBackgroundPatternColor.Blue
  col = getRGBText(R, G, B)
  mat = getPatternName(id)
  
  print('BG Cut = [' + col + '] ' + mat)
  
  sheet.Cells[row,COL_CUTBGPATTERN].Value2 = mat
  sheet.Cells[row,COL_CUTBGCOLOR].Value2   = col
  sheet.Cells[row,COL_CUTBGCOLOR_R].Value2 = str(R)
  sheet.Cells[row,COL_CUTBGCOLOR_G].Value2 = str(G)
  sheet.Cells[row,COL_CUTBGCOLOR_B].Value2 = str(B)
  
  #Shade the background of the colour cell to match the RGB value
  sheet.Cells[row,COL_CUTBGCOLOR].Interior.Color = Color.FromArgb(R,G,B)
  #Determine how dark the colour is and set the font color accordingly
  textcol = _black if (R * 0.299 + G * 0.587 + B * 0.114) > 186 else _white
  sheet.Cells[row,COL_CUTBGCOLOR].Font.Color = textcol
  
  # Get the Foreground Pattern Info
  id = aMaterial.CutForegroundPatternId
  R = aMaterial.CutForegroundPatternColor.Red
  G = aMaterial.CutForegroundPatternColor.Green
  B = aMaterial.CutForegroundPatternColor.Blue
  col = getRGBText(R, G, B)
  mat = getPatternName(id)
  
  print('FG Cut = [' + str(R) + ':' + str(G) + ':' + str(B) + '] ' + mat)
  
  sheet.Cells[row,COL_CUTFGPATTERN].Value2 = mat
  sheet.Cells[row,COL_CUTFGCOLOR].Value2 = col
  sheet.Cells[row,COL_CUTFGCOLOR_R].Value2 = str(R)
  sheet.Cells[row,COL_CUTFGCOLOR_G].Value2 = str(G)
  sheet.Cells[row,COL_CUTFGCOLOR_B].Value2 = str(B)
  
  #Shade the background of the colour cell to match the RGB value
  sheet.Cells[row,COL_CUTFGCOLOR].Interior.Color = Color.FromArgb(R,G,B)
  #Determine how dark the colour is and set the font color accordingly
  textcol = _black if (R * 0.299 + G * 0.587 + B * 0.114) > 186 else _white
  sheet.Cells[row,COL_CUTFGCOLOR].Font.Color = textcol
  
  print('\nSurface Pattern')
  id = aMaterial.SurfaceBackgroundPatternId
  R = aMaterial.SurfaceBackgroundPatternColor.Red
  G = aMaterial.SurfaceBackgroundPatternColor.Green
  B = aMaterial.SurfaceBackgroundPatternColor.Blue
  col = getRGBText(R, G, B)
  mat = getPatternName(id)
  
  print('BG Srf = [' + str(R) + ':' + str(G) + ':' + str(B) + '] ' + mat)
  
  sheet.Cells[row,COL_SURFBGPATTERN].Value2 = mat
  sheet.Cells[row,COL_SURFBGCOLOR].Value2 = col
  sheet.Cells[row,COL_SURFBGCOLOR_R].Value2 = str(R)
  sheet.Cells[row,COL_SURFBGCOLOR_G].Value2 = str(G)
  sheet.Cells[row,COL_SURFBGCOLOR_B].Value2 = str(B)
  
  #Shade the background of the colour cell to match the RGB value
  sheet.Cells[row,COL_SURFBGCOLOR].Interior.Color = Color.FromArgb(R,G,B)
  #Determine how dark the colour is and set the font color accordingly
  textcol = _black if (R * 0.299 + G * 0.587 + B * 0.114) > 186 else _white
  sheet.Cells[row,COL_SURFBGCOLOR].Font.Color = textcol
  
  # Get the Foreground Pattern Info
  id = aMaterial.SurfaceForegroundPatternId
  R = aMaterial.SurfaceForegroundPatternColor.Red
  G = aMaterial.SurfaceForegroundPatternColor.Green
  B = aMaterial.SurfaceForegroundPatternColor.Blue
  col = getRGBText(R, G, B)
  mat = getPatternName(id)
  
  print('FG Srf = [' + str(R) + ':' + str(G) + ':' + str(B) + '] ' + mat)
  
  sheet.Cells[row,COL_SURFFGPATTERN].Value2 = mat
  sheet.Cells[row,COL_SURFFGCOLOR].Value2 = col
  sheet.Cells[row,COL_SURFFGCOLOR_R].Value2 = str(R)
  sheet.Cells[row,COL_SURFFGCOLOR_G].Value2 = str(G)
  sheet.Cells[row,COL_SURFFGCOLOR_B].Value2 = str(B)
  
  #Shade the background of the colour cell to match the RGB value
  sheet.Cells[row,COL_SURFFGCOLOR].Interior.Color = Color.FromArgb(R,G,B)
  #Determine how dark the colour is and set the font color accordingly
  textcol = _black if (R * 0.299 + G * 0.587 + B * 0.114) > 186 else _white
  sheet.Cells[row,COL_SURFFGCOLOR].Font.Color = textcol

# FREEZE THE HEADER ROW
sheet.Application.ActiveWindow.SplitColumn = 0
sheet.Application.ActiveWindow.SplitRow = 1
sheet.Application.ActiveWindow.FreezePanes = True

# CLOSE THE EXCEL DOCUMENT
workbook.SaveAs("S:\\Design Technology\\Code\\Desktop\\Revit\\RevitPythonShell\\DT\\_WIP\\MaterialsToExcel\\TEST.xlsx")
workbook.Close(False)
excel.Quit()
  
  
  