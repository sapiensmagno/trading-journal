VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSysBuilder 
   Caption         =   "System Builder"
   ClientHeight    =   8888
   ClientLeft      =   96
   ClientTop       =   414
   ClientWidth     =   11658
   OleObjectBlob   =   "frmSysBuilder.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmsysbuilder"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit

const stablename as string = "systablefields"
const stablefields as string = "systablefieldscampos"
const stableid as string = "systablefieldsid"
dim beditmode as boolean

private sub userform_initialize()
dim icolumnid as integer
    
    beditmode = false
    
    call sortmultiplecolumns(stablename, "tipo tabela|table")
    call initialize(me)
    icolumnid = application.worksheetfunction.match("#idtablefield", range(stablefields), 0)
    call carrega_treeview(stablename, stablefields, icolumnid, "tipo tabela|table|nome", tvwsysbuilder)
    
end sub


