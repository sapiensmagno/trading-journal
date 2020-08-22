VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportarTrades 
   Caption         =   "Importar arquivo CSV de operações realizadas"
   ClientHeight    =   3696
   ClientLeft      =   48
   ClientTop       =   234
   ClientWidth     =   9336
   OleObjectBlob   =   "frmImportarTrades.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmimportartrades"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
const stablename as string = "tbimportarcsv"
const stablefields as string = "tbimportarcsvcampos"
const stableid as string = "tbimportarcsvid"

dim beditmode as boolean
dim sfile as string

private sub btnabrirarquivo_click()
dim fd as office.filedialog

    set fd = application.filedialog(msofiledialogfilepicker)
    
    with fd
        .filters.clear
        .title = "selecione o arquivo csv"
        .filters.add "csv", "*.csv", 1
        .allowmultiselect = false
        
        if .show = true then
            sfile = .selecteditems(1)
        end if
        
    end with
  
end sub

private sub btnimportar_click()
dim irownro as integer
    
    if trim(sfile) <> "" then
    
        open sfile for input as #1
            irownro = 0
            
            do until eof(1)
                
                line input #1, linefromfile
                lineitems = split(linefromfile, ",")
                msgbox lineitems(0) & lineitems(1) & lineitems(2)
                
                irownro = irownro + 1
                
            loop
            
            
        close #1
        
    end if
end sub

private sub btnok_click()
    unload me
end sub



private sub userform_initialize()
    beditmode = false
    
    'call sortmultiplecolumns(stablename, "categoria|código")
    
    call initialize(me)
end sub

