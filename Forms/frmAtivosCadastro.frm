VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAtivosCadastro 
   Caption         =   "Cadastro de Ativos :: Novo"
   ClientHeight    =   7072
   ClientLeft      =   18
   ClientTop       =   48
   ClientWidth     =   11478
   OleObjectBlob   =   "frmAtivosCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmativoscadastro"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit

const stablename as string = "tbativos"
const stablefields as string = "tbativoscampos"
const stableid as string = "tbativosid"
dim beditmode as boolean

private sub btnedit_click()
    
    beditmode = true
 
    call enabledisablecontrols(me, true)
    
end sub

private sub btnadd_click()
dim obj, totalobj as integer
dim stag as string
dim vtag as variant
    
    call enabledisablecontrols(me, true)
    
    totalobj = me.controls.count - 1
    redim cformatfields(0 to totalobj)
    
    for obj = 0 to totalobj
    
        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if len(trim(me.controls(obj).tag)) > 0 then
        
            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(me.controls(obj).tag, range("systablefields"), 13, 0)
            
            ' sepera as tags em uma lista
            vtag = split(stag, "|")
            
             select case vtag(0)
            
                case "tab"
                
                    me.controls(obj).tabindex = 0
                    
                case "boolean"
                
                    me.controls(obj).value = false
                
                case else
                
                    me.controls(obj).value = ""
                
            end select
        
        end if
        
    next obj ' for obj = 0 to totalobj
end sub

private sub btnsave_click()
dim icolumnid as integer
    
     if len(trim(validacamposobrigatorios(me))) > 0 then
    
        msgbox "preencha os campos: " & chr(10) & chr(13) & validacamposobrigatorios(me)
        exit sub

    else

        call savereg(me, stablename, stablefields, tvwativos, beditmode)
    

        beditmode = false
        
        call enabledisablecontrols(me, false)
        call sortmultiplecolumns(stablename, "categoria|código")
    
        icolumnid = application.worksheetfunction.match("#idativo", range(stablefields), 0)
        call carrega_treeview(stablename, stablefields, icolumnid, "categoria|descrição aux", tvwativos)
    
    end if
    
end sub

private sub btnfechar_click()
    
    unload me
    
end sub

private sub tvwativos_nodeclick(byval node as mscomctllib.node)
dim obj as integer
dim totalobj, icoluna, ilinha as integer
dim stag as string
dim vtable, vtag, vlinha as variant
dim svalor as string

    call mytreeview_findnode(me, node.key)
    
    call enabledisablecontrols(me, false)
    
    totalobj = me.controls.count - 1
    
    for obj = 0 to totalobj
    
        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if trim(me.controls(obj).tag) <> "" then
                
            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(me.controls(obj).tag, range("systablefields"), 13, 0)
                
            ' sepera as tags em uma lista
            vtag = split(stag, "|")
                       
            vtable = split(me.controls(obj).tag, "|")
            icoluna = application.worksheetfunction.match(vtable(1), range(stablefields), 0)
            
            vlinha = split(node.key, "|")
            
            if not isnumeric(vlinha(0)) then exit sub
            
            ilinha = vlinha(0)

            select case vtag(0)
                
                case "radio"
                
                    if application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna) = me.controls(obj).caption then
                        
                        me.controls(obj).value = true
                        
                    end if
                    
                case "value", "money", "number"
                
                    me.controls(obj).value = format(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), vtag(2))
                
                case else
                    
                    me.controls(obj).value = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
    
        next obj ' for obj = 0 to totalobj
end sub

private sub txtativonome_exit(byval cancel as msforms.returnboolean)

    if trim(txtativocodigo.value) <> "" and trim(txtativonome.value) <> "" then
        
        txtativodescricao.value = txtativocodigo & " - " & txtativonome.value
    
    end if
    
end sub

private sub userform_initialize()
dim icolumnid as integer
    
    beditmode = false
    call enabledisablecontrols(me, false)
    
    call sortmultiplecolumns(stablename, "categoria|código")
    
    call initialize(me)
    icolumnid = application.worksheetfunction.match("#idativo", range(stablefields), 0)
    call carrega_treeview(stablename, stablefields, icolumnid, "categoria|descrição aux", tvwativos)
    
end sub

