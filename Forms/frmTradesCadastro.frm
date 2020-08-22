VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTradesCadastro 
   Caption         =   "Registro de trade :: Novo Registro"
   ClientHeight    =   7938
   ClientLeft      =   -150
   ClientTop       =   -642
   ClientWidth     =   15042
   OleObjectBlob   =   "frmTradesCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
end
attribute vb_name = "frmtradescadastro"
attribute vb_globalnamespace = false
attribute vb_creatable = false
attribute vb_predeclaredid = true
attribute vb_exposed = false
option explicit

dim filename as variant
dim newfilepath01, newfilepath02 as string
dim dir as string
dim beditmode as boolean
dim objtrade as ctrades

const stablename as string = "tbtrades"
const stablefields as string = "tbtradescampos"
const stableid as string = "tbtradesid"

private sub carregatrades()
dim ilinha as integer
 
    ilinha = 2
    do until sheets("tbtrades").cells(ilinha, 1) = ""
        cbxtrade.additem sheets("tbtrades").cells(ilinha, 1)
        ilinha = ilinha + 1
    loop

end sub

private function formatfields(byval control as object) as string
    dim stag as string
    dim vtag as variant

    stag = application.worksheetfunction.vlookup(control.tag, range("systablefields"), 13, 0)
    vtag = split(stag, "|")

    formatfields = format(control.value, vtag(2))
end function

private sub btnhoraentrada_click()

    txtdataentrada.value = format(now(), "dd/mm/yyyy")
    txthoraentrada.value = format(now(), "hh:mm:ss")

end sub

private sub btnhorasaida1_click()

    txthorasaida1.value = format(now(), "hh:mm:ss")
    
end sub

private sub btnhorasaida2_click()

    txthorasaida2.value = format(now(), "hh:mm:ss")
    
end sub

private sub btnhorasaida3_click()

    txthorasaida3.value = format(now(), "hh:mm:ss")
    
end sub

private sub btnsalvar_click()

    dim llastregid, llastrow, lnewregid, ilinha as long

    set objtrade = new ctrades

    with objtrade
        .dataentrada = format(now(), "dd/mm/yyyy")
    end with
    
    set objtrade = nothing
    
    
    if not beditmode then
        
        lnewregid = clng(txtnrotrade.value)
            
        llastrow = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).row
        llastregid = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).value
        lnewregid = llastregid + 1

        if llastrow = 2 and llastregid = empty then
            ilinha = 2
        else
            ilinha = llastrow + 1
        end if
    else
        
        lnewregid = clng(cbxtrade.value)
        
        ilinha = application.worksheetfunction.match(cint(cbxtrade.value), range(stableid), 0)
    
    end if

    call salvar(me, stablename, stablefields, ilinha, lnewregid, beditmode, lnewregid, cstr(newfilepath01), cstr(newfilepath02))
    activeworkbook.save
    call refreshpowequery
    
 
end sub

private sub btnfechar_click()
        
    unload me
    
end sub

private sub btnviewimagem01_click()
    
    frmviewimage.imgviewimage.picture = imgscreenshot01.picture
    frmviewimage.repaint
    frmviewimage.show
    
end sub

private sub btnviewimagem02_click()

    frmviewimage.imgviewimage.picture = imgscreenshot02.picture
    frmviewimage.repaint
    frmviewimage.show

end sub

private sub copiarprocessodecisorio(byval sativo as string)
dim icoldatasaida, icolativo, icolcondicaomercado, icolriscosempotencial as integer
dim scondicaomercado, sriscosempotencial, sativoreg as string
dim ddatelasttrade, dlastdatereg as date
dim llastrow as long

    if trim(sativo) = "" then exit sub

    icoldatasaida = application.worksheetfunction.match("data saída", range("tbtradescampos"), 0)
    llastrow = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, icoldatasaida).end(xlup).offset(0, 0).row
    ddatelasttrade = application.worksheetfunction.index(range("tbtradesdatasaida"), llastrow)
    
    while (cdate(format(now(), "dd/mm/yyyy")) = cdate(ddatelasttrade))
    
        sativoreg = application.worksheetfunction.index(range("tbtradesativo"), llastrow)
    
        if sativo = sativoreg then
        
            icolcondicaomercado = application.worksheetfunction.match("condição mercado", range("tbtradescampos"), 0)
            scondicaomercado = application.worksheetfunction.index(range("tbtradescondicaomercado"), llastrow)
    
            call carregachecklists(scondicaomercado, lbxcondicaomercado)
    
            icolriscosempotencial = application.worksheetfunction.match("riscos em potencial", range("tbtradescampos"), 0)
            sriscosempotencial = application.worksheetfunction.index(range("tbtradesriscosempotencial"), llastrow)
    
            call carregachecklists(sriscosempotencial, lbxriscosempotencial)
    
        end if
        
        llastrow = llastrow - 1
        ddatelasttrade = application.worksheetfunction.index(range("tbtradesdatasaida"), llastrow)
    
    wend

end sub

private sub cbxativo_exit(byval cancel as msforms.returnboolean)
dim sconfigkey as string
dim ilinha as long

    if len(trim(cbxcorretora.value)) > 0 and len(trim(cbxativo.value)) > 0 then
        sconfigkey = cbxcorretora.value & "_" & cbxativo.value
        
        ilinha = application.worksheetfunction.match(sconfigkey, range("tbconfiguracaoesgeraissearchkey"), 0)
        
        if trim(cbxestrategia.value) = "" then
            cbxestrategia.value = application.worksheetfunction.index(range("tbconfiguracoesgeraisestrategia"), ilinha)
        end if
        
        if trim(cbxtipoentrada.value) = "" then
            cbxtipoentrada.value = application.worksheetfunction.index(range("tbconfiguracoesgeraistipodeentrada"), ilinha)
        end if
        
        if trim(cbxtipotrade.value) = "" then
            cbxtipotrade.value = application.worksheetfunction.index(range("tbconfiguracoesgeraistipodeoperacao"), ilinha)
        end if
        
        if trim(cbxtipoconta.value) = "" then
            cbxtipoconta.value = application.worksheetfunction.index(range("tbconfiguracoesgeraistipodeconta"), ilinha)
        end if
        
        if trim(txtmotivacaoentrada.value) = "" then
            txtmotivacaoentrada.value = application.worksheetfunction.index(range("tbconfiguracoesgeraismotivacaoentrada"), ilinha)
        end if
        
        if trim(txtmotivacaosaida.value) = "" then
            txtmotivacaosaida.value = application.worksheetfunction.index(range("tbconfiguracoesgeraismotivacaosaida"), ilinha)
        end if
        
    end if
    
    call copiarprocessodecisorio(trim(cbxativo.value))

end sub

private sub cbxtrade_change()
dim obj as integer
dim totalobj, icoluna, ilinha as integer
dim stag as string
dim vtable, vtag as variant
dim svalor as string
        
    beditmode = true
    
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
            if trim(cbxtrade.value) = "" and beditmode then exit sub
            ilinha = clng(cbxtrade.value)

            select case vtag(0)
            
                case "date", "time"
                
                     me.controls(obj).value = format(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), vtag(2))
                
                case "checklist" ' trata os campos checklist
                
                    dim vlistaitens as variant
                    dim icountitens, x, y, z as integer
                
                    vlistaitens = split(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), ";")
                    z = ubound(vlistaitens)
                    icountitens = me.controls(obj).listcount
                    
                    ' limpa todos os itens selecionados
                    for x = 0 to icountitens
                    
                        me.controls(obj).selected(x) = false
                        
                    next x
                    
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
                
                case "imagem"
                
                    me.controls(obj).picture = loadpicture(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna))
                    
                    if vtable(1) = "imagem #01" then
                        newfilepath01 = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                    elseif vtable(1) = "imagem #02" then
                        newfilepath02 = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                    end if
                    
                case "calculado"
                    
                    me.controls(obj).value = format(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), vtag(2))
                    
                case "value", "money"
                
                    me.controls(obj).value = format(application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna), vtag(2))
                
                case else
                    
                    me.controls(obj).value = application.worksheetfunction.vlookup(ilinha, range(stablename), icoluna)
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
    
        next obj ' for obj = 0 to totalobj

end sub

private sub imgscreenshot02_click()
  dim finfo as string
  dim filterindex as integer
  dim title as string
  dim file as variant
  dim filename as string
  dim filepath as string
  
  'application.defaultfilepath
  
  finfo = "imagens (*.jpg*),*.*"
  filterindex = 1
  title = "selecione um arquivo para importar"
  file = application.getopenfilename(finfo, filterindex, title)

  if file = false then exit sub
  
  filepath = file
  filename = right(filepath, len(filepath) - instrrev(filepath, "\"))
  newfilepath02 = filepath
  'thisworkbook.path & "\imagens\" & filename

  on error goto erro
    imgscreenshot02.picture = loadpicture(filepath)
    
    frmtradescadastro.repaint
    exit sub
    
erro:
     msgbox "não foi possível carregar a imagem" & err.number & err.description
  
  imgscreenshot02_click
    
end sub

private sub imgscreenshot01_click()
  dim finfo as string
  dim filterindex as integer
  dim title as string
  dim file as variant
  dim filename as string
  dim filepath as string
  
  'application.defaultfilepath
  
  finfo = "imagens (*.jpg*),*.*"
  filterindex = 1
  title = "selecione um arquivo para importar"
  file = application.getopenfilename(finfo, filterindex, title)

  if file = false then exit sub
  
  filepath = file
  filename = right(filepath, len(filepath) - instrrev(filepath, "\"))
  newfilepath01 = filepath
  'thisworkbook.path & "\imagens\" & filename

  on error goto erro
    imgscreenshot01.picture = loadpicture(filepath)
    
    frmtradescadastro.repaint
    exit sub
    
erro:
     msgbox "não foi possível carregar a imagem" & err.number & err.description
  
  imgscreenshot01_click
    
end sub

private sub lblativo_click()
    
    frmativoscadastro.show
    
end sub

private sub lblcorretora_click()
    
    frmcorretorascadastro.show
    
end sub

private sub lblestrategia_click()

    frmestrategiascadastro.show
    
end sub

private sub txtdataentrada_keypress(byval keyascii as msforms.returninteger)
dim strvalid as string

    strvalid = "0123456789"
    
    if instr(strvalid, chr(keyascii)) = 0 then
        keyascii = 0
    else
        if txtdataentrada.selstart = 2 then
            txtdataentrada.seltext = "/"
        elseif txtdataentrada.selstart = 5 then
            txtdataentrada.seltext = "/"
        end if
    end if
end sub

private sub txtdatasaida_keypress(byval keyascii as msforms.returninteger)
dim strvalid as string

    strvalid = "0123456789"
    
    if instr(strvalid, chr(keyascii)) = 0 then
        keyascii = 0
    else
        if txtdatasaida.selstart = 2 then
            txtdatasaida.seltext = "/"
        elseif txtdatasaida.selstart = 5 then
            txtdatasaida.seltext = "/"
        end if
    end if
end sub

private sub txthoraentrada_keypress(byval keyascii as msforms.returninteger)
dim strvalid as string

    strvalid = "0123456789"
    
    if instr(strvalid, chr(keyascii)) = 0 then
        keyascii = 0
    else
        if txthoraentrada.selstart = 2 then
            txthoraentrada.seltext = ":"
        elseif txthoraentrada.selstart = 5 then
            txthoraentrada.seltext = ":"
        elseif txthoraentrada.selstart = 8 then
            txthoraentrada.seltext = ":"
        end if
    end if
end sub

private sub txthorasaida_keypress(byval keyascii as msforms.returninteger)
dim strvalid as string

    strvalid = "0123456789"
    
    if instr(strvalid, chr(keyascii)) = 0 then
        keyascii = 0
    else
        if txthorasaida.selstart = 2 then
            txthorasaida.seltext = ":"
        elseif txthorasaida.selstart = 5 then
            txthorasaida.seltext = ":"
        elseif txthorasaida.selstart = 8 then
            txthorasaida.seltext = ":"
        end if
    end if
end sub

private sub txtlucroprejuizo_exit(byval cancel as msforms.returnboolean)

    txtlucroprejuizo.value = formatfields(txtlucroprejuizo)

end sub

private sub txtmen_exit(byval cancel as msforms.returnboolean)

    txtmen.value = formatfields(txtmen)

end sub

private sub txtmep_exit(byval cancel as msforms.returnboolean)

    txtmep.value = formatfields(txtmep)

end sub

private sub txtnrocontratos_exit(byval cancel as msforms.returnboolean)

    txtnrocontratos.value = formatfields(txtnrocontratos)

end sub

private sub txtprecosaida_exit(byval cancel as msforms.returnboolean)
    dim vvalortick, vqtdpontos as double
    
    if (optcompra.value) then
        txtpontos.value = (txtprecosaida.value - txtprecoentrada.value)
    elseif (optvenda.value) then
        txtpontos.value = (txtprecoentrada.value - txtprecosaida.value)
    end if
    
    exit sub
    txtprecosaida.value = formatfields(txtprecosaida)
    
    if trim(txtprecosaida.value) <> "" and trim(txtprecoentrada.value) then

        vvalortick = application.worksheetfunction.index(range("tbativosvalortick"), application.worksheetfunction.match(cbxativo.value, range("tbativosdescriçãoaux"), 0))
        if optcompra.value then

            vqtdpontos = txtprecosaida.value - txtprecoentrada.value

        elseif optvenda.value then

            vqtdpontos = txtprecoentrada.value - txtprecosaida.value

        end if

        txtlucroprejuizo.value = (vqtdpontos * vvalortick) * cint(txtnrocontratos.value)
    end if

end sub
private sub txtprecosaida2_exit(byval cancel as msforms.returnboolean)
    if (len(txtpontos2.value) and len(txtprecosaida2.value) and len(txtprecoentrada.value)) then
        if (optcompra.value) then
             txtpontos2.value = (txtprecosaida2.value - txtprecoentrada.value)
         elseif (optvenda.value) then
             txtpontos2.value = (txtprecoentrada.value - txtprecosaida2.value)
         end if
    else
        txtpontos2.value = ""
    end if
end sub
private sub txtprecosaida3_exit(byval cancel as msforms.returnboolean)
    if (len(txtpontos3.value) and len(txtprecosaida3.value) and len(txtprecoentrada.value)) then
        if (optcompra.value) then
            txtpontos3.value = (txtprecosaida3.value - txtprecoentrada.value)
        elseif (optvenda.value) then
            txtpontos3.value = (txtprecoentrada.value - txtprecosaida3.value)
        end if
    else
        txtpontos3.value = ""
    end if
end sub

private sub txtprecosl_exit(byval cancel as msforms.returnboolean)

    txtprecosl.value = formatfields(txtprecosl)

end sub

private sub txtprecoentrada_exit(byval cancel as msforms.returnboolean)

    txtprecoentrada.value = formatfields(txtprecoentrada)

end sub

private sub txtprecotp_exit(byval cancel as msforms.returnboolean)

    txtprecotp.value = formatfields(txtprecotp)

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

private sub userform_initialize()
dim llastrow, llastregid, lnewregid as long

    beditmode = false
    
    llastrow = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).row
    llastregid = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).value
    
    if llastregid = empty then
        lnewregid = 1
    else
        lnewregid = llastregid + 1
    end if
        
    txtnrotrade.value = lnewregid
    
    call initialize(me)
    carregatrades
    
end sub

