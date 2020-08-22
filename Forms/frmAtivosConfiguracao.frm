VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAtivosConfiguracao 
   Caption         =   "Configuração de Ativos :: Novo Registro"
   ClientHeight    =   7552
   ClientLeft      =   96
   ClientTop       =   414
   ClientWidth     =   11634
   OleObjectBlob   =   "frmAtivosConfiguracao.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmativosconfiguracao"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit

const stablename as string = "tbativoconfiguracoes"
const stablefields as string = "tbativoconfiguracoescampos"
const stableid as string = "tbativoconfiguracoesid"
dim beditmode as boolean

sub mytreeview_findnode(byval strkey as string)
dim mynode as node

    for each mynode in me.tvwativos.nodes
   
        mynode.bold = mynode.key = strkey
        
    next
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

private sub btnedit_click()

    beditmode = true

    call enabledisablecontrols(me, true)

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
    
        call savereg(me, stablename, stablefields, tvwativos, beditmode)
            
        beditmode = false
        
        call enabledisablecontrols(me, false)
        call sortmultiplecolumns(stablename, "categoria|ativo")
        icolumnid = application.worksheetfunction.match("#idativoconfiguracao", range(stablefields), 0)
        call carrega_treeview(stablename, stablefields, icolumnid, "categoria|ativo|corretora", tvwativos)

    end if
    
end sub

private sub tvwativos_nodeclick(byval node as mscomctllib.node)
dim obj as integer
dim totalobj, icoluna, ilinha as integer
dim stag as string
dim vtable, vtag, vlinha as variant
dim svalor as string

    call mytreeview_findnode(node.key)
    
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
                    
                case "value", "money", "number"
                
                    me.controls(obj).value = format(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), vtag(2))
                
                case else
                    
                    me.controls(obj).value = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
    
        next obj ' for obj = 0 to totalobj
end sub

private sub userform_initialize()
dim icolumnid as integer

    beditmode = false
    call enabledisablecontrols(me, false)
    
    call sortmultiplecolumns(stablename, "categoria|ativo|corretora")

    call initialize(me)
    icolumnid = application.worksheetfunction.match("#idativoconfiguracao", range(stablefields), 0)
    call carrega_treeview(stablename, stablefields, icolumnid, "categoria|ativo|corretora", tvwativos)
end sub


