VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
option explicit

private sub worksheet_activate()

    activewindow.displayheadings = false
    activewindow.displaygridlines = false
    application.displayformulabar = false

end sub

private sub worksheet_deactivate()

    activewindow.displayheadings = true
    activewindow.displaygridlines = true
    application.displayformulabar = true

end sub


