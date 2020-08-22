VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChecklistDiario 
   Caption         =   "CheckList Diário"
   ClientHeight    =   7088
   ClientLeft      =   96
   ClientTop       =   414
   ClientWidth     =   11118
   OleObjectBlob   =   "frmChecklistDiario.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmchecklistdiario"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit


const stablename as string = "tbchecklistdiario"
const stablefields as string = "tbchecklistdiariocampos"
const stableid as string = "tbchecklistdiarioid"

dim beditmode as boolean

private sub btnsalvar_click()
dim obj, totalobj, icoluna, ilinha as integer
dim llastrow, lnewregid, llastregid as long
dim vtag, vtable, vlinha as variant
dim stag as string
dim dlastdate as date
      
    llastrow = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).row
    llastregid = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).value
    lnewregid = llastregid + 1
    
    dlastdate = application.worksheetfunction.index(range("tbchecklistdiariodata"), llastrow)
    
    if (cdate(format(now(), "dd/mm/yyyy")) = cdate(dlastdate)) then
    
        ilinha = llastrow
        
    else
        if llastrow = 2 and llastregid = empty then
            ilinha = 2
        else
            ilinha = llastrow + 1
        end if
    
    end if

    totalobj = me.controls.count - 1
                
    'worksheets(stablename).cells(ilinha, 1) = ilinha 'application.worksheetfunction.match(ilinha, range(stableid), 1)
        
    for obj = 0 to totalobj
    
        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if trim(me.controls(obj).tag) <> "" then
            
            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(me.controls(obj).tag, range("systablefields"), 13, 0)
            
            ' sepera as tags em uma lista
            vtag = split(stag, "|")
            
            vtable = split(me.controls(obj).tag, "|")
            icoluna = application.worksheetfunction.match(vtable(1), range(stablefields), 0)
            
            'range(stablename).rows.count + 1

            select case vtag(0)
            
                case "data"
                
                    worksheets(stablename).cells(ilinha, icoluna) = cdate(format(now(), "dd/mm/yyyy"))
                    
                case "radio" ' trata os campos tipo optionbutton

                        if me.controls(obj).value then
                            worksheets(stablename).cells(ilinha, icoluna) = me.controls(obj).caption
                        end if
                
                case "checklist" ' trata os campos checklist

                    dim spicklist, spicklist2 as string
                    dim i as integer

                    spicklist = ""
                    spicklist2 = ""
                    for i = 0 to me.controls(obj).listcount - 1

                       if me.controls(obj).selected(i) = true then

                          if spicklist = "" then
                             spicklist = me.controls(obj).list(i)
                           else
                             spicklist2 = me.controls(obj).list(i)
                             spicklist = spicklist & ";" & spicklist2
                           end if

                        end if

                    next

                    worksheets(stablename).cells(ilinha, icoluna) = spicklist
                    
                case else
                    
                    worksheets(stablename).cells(ilinha, icoluna) = me.controls(obj).value
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
        
        next obj ' for obj = 0 to totalobj
        
        unload me
end sub

public sub initialize()
dim obj as integer
dim totalobj as integer
dim vtag as variant
dim stag as string
        
    totalobj = me.controls.count - 1
    redim cformatfields(0 to totalobj)
    
    for obj = 0 to totalobj
    
        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if trim(me.controls(obj).tag) <> "" then
                
            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(me.controls(obj).tag, range("systablefields"), 13, 0)
            
            ' sepera as tags em uma lista
            vtag = split(stag, "|")

            select case vtag(0)
                    
                case "combo", "checklist" ' carrega a lista de itens da tabela lookup na lista
                
                    dim rtablesource, rtablelookup as range
                    dim sfieldlookup as string
                    dim ifieldlist, ifieldordem as integer
                    dim linha as integer
                    dim sstringaux as string

                    set rtablesource = range(vtag(4))
                    set rtablelookup = range(vtag(5))
                    sfieldlookup = vtag(6)
                    ifieldlist = application.worksheetfunction.match(sfieldlookup, rtablelookup, 0)
                    ifieldordem = application.worksheetfunction.match("ordem", rtablelookup, 0)

                    linha = 1
                    sstringaux = ""
                    me.controls(obj).clear
                    do until rtablesource.cells(linha, ifieldlist) = ""
                        
                        if sstringaux <> rtablesource.cells(linha, ifieldlist) then
                            
                            if me.controls(obj).name = "lbxchecklistitemsabertura" and rtablesource.cells(linha, ifieldordem) = "abertura" then
                                me.controls(obj).additem rtablesource.cells(linha, ifieldlist)
                                sstringaux = rtablesource.cells(linha, ifieldlist)
                            elseif me.controls(obj).name = "lbxchecklistitemsfechamento" and rtablesource.cells(linha, ifieldordem) = "fechamento" then
                                me.controls(obj).additem rtablesource.cells(linha, ifieldlist)
                                sstringaux = rtablesource.cells(linha, ifieldlist)
                            end if
                        end if
                        
                        linha = linha + 1
                    
                    loop
                    
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
        
        next obj ' for obj = 0 to totalobj
end sub

private sub lbxchecklistitemsabertura_change()
dim sitemselecionado as string
dim par2 as integer

    sitemselecionado = lbxchecklistitemsabertura.list(lbxchecklistitemsabertura.listindex)
   
    par2 = application.worksheetfunction.match(sitemselecionado, range("tbchecklistitem"), 0)
    txtdescricaoitemabertura.value = application.worksheetfunction.index(range("tbchecklistdescrição"), par2)
end sub

private sub lbxchecklistitemsfechamento_change()
dim sitemselecionado as string
dim par2 as integer

    sitemselecionado = lbxchecklistitemsfechamento.list(lbxchecklistitemsfechamento.listindex)
   
    par2 = application.worksheetfunction.match(sitemselecionado, range("tbchecklistitem"), 0)
    txtdescricaoitemfechamento.value = application.worksheetfunction.index(range("tbchecklistdescrição"), par2)
end sub

private sub userform_initialize()
            
    txtdata.value = format(now(), "dd/mm/yyyy")
        
    beditmode = true
    call initialize
    call carregaitens
    
end sub

private sub carregaitens()
dim llastrow as long
dim icoldata as integer
dim dlastdate as date
dim sitensabertura, sitensfechamento as string

    llastrow = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).row
    dlastdate = application.worksheetfunction.index(range("tbchecklistdiariodata"), llastrow)
    
    if (cdate(format(now(), "dd/mm/yyyy")) = cdate(dlastdate)) then
    
        sitensabertura = application.worksheetfunction.index(range("tbchecklistdiarioitensabertura"), llastrow)
        call carregachecklists(sitensabertura, lbxchecklistitemsabertura)
        
        txtobservacoesabertura.value = application.worksheetfunction.index(range("tbchecklistdiarioobsabertura"), llastrow)
        
        sitensfechamento = application.worksheetfunction.index(range("tbchecklistdiarioitensfechamento"), llastrow)
        call carregachecklists(sitensfechamento, lbxchecklistitemsfechamento)
        
        txtobservacoesfechamento.value = application.worksheetfunction.index(range("tbchecklistdiarioobsfechamento"), llastrow)
        
        call carregaradio(llastrow)
        
    end if
    
end sub

private sub carregachecklists(byval sstringitens as string, byref ome as object)
dim vlistaitens as variant
dim icountitens, x, y, z, obj as integer
    
    vlistaitens = split(sstringitens, ";")
    z = ubound(vlistaitens)
    icountitens = ome.listcount
    
    ' limpa todos os itens selecionados
    for x = 0 to icountitens
    
        ome.selected(x) = false
        
    next x
    
    y = 0
    for x = 0 to icountitens
    
        if (y <= z) then
            if (vlistaitens(y) = ome.list(x)) then
                ome.selected(x) = true
                                   
                y = y + 1
            end if
        end if
        
    next x
end sub

private sub carregaradio(byval ilinha as integer)
dim obj as integer
dim totalobj, icoluna as integer
dim stag as string
dim vtable, vtag, vlinha as variant
dim svalor as string
    
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
            
            select case vtag(0)
                
                case "radio"
                
                    if application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna) = me.controls(obj).caption then
                        
                        me.controls(obj).value = true
                        
                    end if
                    
               
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
    
        next obj ' for obj = 0 to totalobj

end sub

