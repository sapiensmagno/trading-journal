VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorretorasCadastro 
   Caption         =   "Cadastro de Corretoras :: Novo Registro"
   ClientHeight    =   7080
   ClientLeft      =   -66
   ClientTop       =   -288
   ClientWidth     =   11412
   OleObjectBlob   =   "frmCorretorasCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmcorretorascadastro"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit

const stablename as string = "tbcorretoras"
const stablefields as string = "tbcorretorascampos"
const stableid as string = "tbcorretorasid"
dim beditmode as boolean

sub mytreeview_findnode(byval strkey as string)
dim mynode as node

    for each mynode in me.tvwcorretoras.nodes
   
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

private sub btnexibir_click()

    if btnexibir.value then
    
        txtsenha.passwordchar = ""
        txtassinaturaeletronica.passwordchar = ""
    
    else
    
        txtsenha.passwordchar = "*"
        txtassinaturaeletronica.passwordchar = "*"
    
    end if
end sub

private sub btnfechar_click()

    unload me
    
end sub

private sub btnedit_click()
    
    beditmode = true
 
    call enabledisablecontrols(me, true)
    
end sub

private sub btnsave_click()
dim icolumnid as integer
    
     if len(trim(validacamposobrigatorios(me))) > 0 then
    
        msgbox "preencha os campos: " & chr(10) & chr(13) & validacamposobrigatorios(me)
        exit sub

    else
    
        call savereg(me, stablename, stablefields, tvwcorretoras, beditmode)
            
        beditmode = false
        
        call enabledisablecontrols(me, false)
        call sortmultiplecolumns(stablename, "nacional/internacional|nome")
        icolumnid = application.worksheetfunction.match("#idcorretora", range(stablefields), 0)
        call carrega_treeview(stablename, stablefields, icolumnid, "nacional/internacional|nome", tvwcorretoras)
    end if
    
end sub

private sub tvwcorretoras_nodeclick(byval node as mscomctllib.node)
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
    
    call sortmultiplecolumns(stablename, "nacional/internacional|nome")

    call initialize(me)
    icolumnid = application.worksheetfunction.match("#idcorretora", range(stablefields), 0)
    call carrega_treeview(stablename, stablefields, icolumnid, "nacional/internacional|nome", tvwcorretoras)

end sub





