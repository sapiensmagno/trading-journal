VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEstrategiasCadastro 
   Caption         =   "Cadastro de Estratégias :: Novo Registro"
   ClientHeight    =   8320
   ClientLeft      =   18
   ClientTop       =   66
   ClientWidth     =   12870
   OleObjectBlob   =   "frmEstrategiasCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmestrategiascadastro"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit

const stablename as string = "tbestrategias"
const stablefields as string = "tbestrategiascampos"
const stableid as string = "tbestrategiasid"

dim beditmode as boolean

private sub enabledisablecontrols(byval benabled as boolean)
dim obj, totalobj as integer
dim stag as string
dim vtag as variant

const lenablecolor as long = &h80000005
const ldesablecolor as long = &h8000000f
    
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
            
                case "tab", "boolean"
                
                    if benabled then
            
                        me.controls(obj).enabled = benabled
                        me.controls(obj).backcolor = &h8000000f
                        
                    else
                    
                        me.controls(obj).enabled = false
                    
                    end if
                
                case else
                
                    if benabled then
            
                        me.controls(obj).enabled = benabled
                        me.controls(obj).backcolor = lenablecolor
                        
                    else
                    
                        me.controls(obj).enabled = false
                        me.controls(obj).backcolor = ldesablecolor
                    
                    end if
                
            end select
        
        end if
        
    next obj ' for obj = 0 to totalobj
end sub

sub mytreeview_findnode(byval strkey as string)
dim mynode as node

    for each mynode in me.tvwestrategias.nodes
   
        mynode.bold = mynode.key = strkey
        
    next
end sub

private sub btnadd_click()
dim obj, totalobj as integer
dim stag as string
dim vtag as variant
    
    call enabledisablecontrols(true)
    
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

private sub btnedit_click()
dim obj as integer
dim totalobj, icoluna, ilinha as integer
dim stag as string
dim vtable, vtag, vlinha as variant
dim svalor as string
        
    beditmode = true
    
    call enabledisablecontrols(true)
    
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
            
            vlinha = split(tvwestrategias.selecteditem.key, "|")
            
            if not isnumeric(vlinha(0)) then exit sub
            
            ilinha = vlinha(0)

            select case vtag(0)
                
                case "checklist" ' trata os campos checklist
                
                    dim vlistaitens as variant
                    dim icountitens, x, y, z as integer
                
                    vlistaitens = split(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), ";")
                    z = ubound(vlistaitens)
                    icountitens = me.controls(obj).listcount
                    
                    y = 0
                    for x = 0 to icountitens
                    
                        if (y <= z) then
                            if (vlistaitens(y) = me.controls(obj).list(x)) then
                                me.controls(obj).selected(x) = true
                                                   
                                y = y + 1
                            end if
                        end if
                        
                    next x
                
                case else
                    
                    me.controls(obj).value = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
    
        next obj ' for obj = 0 to totalobj
    
end sub

private sub btnfechar_click()
    
    unload me
    
end sub

private sub btnsave_click()
    dim icolumnid as integer

    if len(trim(validacamposobrigatorios(me))) > 0 then

        msgbox "preencha os campos: " & chr(10) & chr(13) & validacamposobrigatorios(me)
        exit sub

    else

        call savereg(me, stablename, stablefields, tvwestrategias, beditmode)
            
        beditmode = false
        
        call enabledisablecontrols(false)
        call sortmultiplecolumns(stablename, "categoria ativo|nome")
    
        icolumnid = application.worksheetfunction.match("#idestrategia", range(stablefields), 0)
        call carrega_treeview(stablename, stablefields, icolumnid, "categoria ativo|nome", tvwestrategias)
        
    end if
    
end sub

private sub tvwestrategias_nodeclick(byval node as mscomctllib.node)
dim obj as integer
dim totalobj, icoluna, ilinha as integer
dim stag as string
dim vtable, vtag, vlinha as variant
dim svalor as string

    call mytreeview_findnode(node.key)
    
    call enabledisablecontrols(false)
    
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
            
                case "date", "time"
                
                     me.controls(obj).value = format(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), vtag(2))
                
                case "checklist" ' trata os campos checklist
                
                    dim vlistaitens as variant
                    dim icountitens, x, y, z as integer
                
                    vlistaitens = split(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), ";")
                    z = ubound(vlistaitens)
                    icountitens = me.controls(obj).listcount
                    
                    y = 0
                    for x = 0 to icountitens
                    
                        if (y <= z) then
                            if (vlistaitens(y) = me.controls(obj).list(x)) then
                                me.controls(obj).selected(x) = true
                                                   
                                y = y + 1
                            end if
                        end if
                        
                    next x
                
                case "radio"
                
                    if application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna) = me.controls(obj).caption then
                        
                        me.controls(obj).value = true
                        
                    end if
                
                case else
                    
                    me.controls(obj).value = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
    
        next obj ' for obj = 0 to totalobj
end sub

private sub userform_initialize()
dim icolumnid as integer

    beditmode = false
    call enabledisablecontrols(false)
    
    call sortmultiplecolumns(stablename, "categoria ativo|nome")

    call initialize(me)
    icolumnid = application.worksheetfunction.match("#idestrategia", range(stablefields), 0)
    call carrega_treeview(stablename, stablefields, icolumnid, "categoria ativo|nome", tvwestrategias)

end sub

