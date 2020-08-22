Attribute VB_Name = "mGlobalEvents"
Option Explicit
Dim cFormatFields() As New cFormatFields

Public Sub Salvar(ByRef oMe As Object, ByVal sTableName As String, _
                  ByVal sTableFields As String, ByVal iLinha As Long, _
                  ByVal lNewRegId As Long, ByVal bEditMode As Boolean, _
                  ByVal lTradeId As Long, ByVal newFilePath01 As String, _
                  ByVal newFilePath02 As String)
dim obj as integer
dim totalobj, icoluna as integer
dim vtag, vtable as variant
dim stag as string
dim smsgpersistir as string


    if len(trim(validacamposobrigatorios(ome))) > 0 then
    
        msgbox "preencha os campos: " & chr(10) & chr(13) & validacamposobrigatorios(ome)

    else
       
        totalobj = ome.controls.count - 1
        'redim cformatfields(0 to totalobj)
        
        for obj = 0 to totalobj
        
            ' verifica se a propriedade tag tem conteúdo e se o campo é físico
            if trim(ome.controls(obj).tag) <> "" then
                
                ' retorna a tag de metadados do campo na tabela de configurações
                stag = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 13, 0)
                
                ' sepera as tags em uma lista
                vtag = split(stag, "|")
                
                vtable = split(ome.controls(obj).tag, "|")
                icoluna = application.worksheetfunction.match(vtable(1), range(stablefields), 0)
                
                'range(stablename).rows.count + 1

                select case vtag(0)
                
                    case "id"
                        if not beditmode then
                    
                            worksheets(stablename).cells(ilinha, icoluna) = lnewregid
                        
                        else
                        
                            worksheets(stablename).cells(ilinha, icoluna) = ltradeid
                        
                        end if
                
                    case "date", "time"
                        dim ddate as date
                        
                        ddate = iif(len(ome.controls(obj).value), ome.controls(obj).value, empty)
                        
                    
                        worksheets(stablename).cells(ilinha, icoluna) = ddate
                    
                    case "checklist" ' trata os campos checklist
                    
                        dim spicklist, spicklist2 as string
                        dim i as integer
                        
                        spicklist = ""
                        spicklist2 = ""
                        for i = 0 to ome.controls(obj).listcount - 1
                        
                           if ome.controls(obj).selected(i) = true then
                           
                              if spicklist = "" then
                                 spicklist = ome.controls(obj).list(i)
                               else
                                 spicklist2 = ome.controls(obj).list(i)
                                 spicklist = spicklist & ";" & spicklist2
                               end if
                               
                            end if
                            
                        next
                        
                        worksheets(stablename).cells(ilinha, icoluna) = spicklist
                    
                    case "radio" ' trata os campos tipo optionbutton

                        if ome.controls(obj).value then
                            worksheets(stablename).cells(ilinha, icoluna) = ome.controls(obj).caption
                        end if
                    
                    case "imagem"
                    
                        if vtable(1) = "imagem #01" then
                            worksheets(stablename).cells(ilinha, icoluna) = newfilepath01
                        elseif vtable(1) = "imagem #02" then
                            worksheets(stablename).cells(ilinha, icoluna) = newfilepath02
                        end if
                        
                    case "money"
                        dim cmoney as currency
                        
                        cmoney = ome.controls(obj).value
                        
                        worksheets(stablename).cells(ilinha, icoluna) = cmoney
                    
                    case "value"
                        dim cvalue as double
                        
                        cvalue = iif(len(ome.controls(obj).value), ome.controls(obj).value, 0)
                        
                        worksheets(stablename).cells(ilinha, icoluna) = cvalue
                        
                    case "calculado"
                    
                    
                    case else
                        
                        worksheets(stablename).cells(ilinha, icoluna) = ome.controls(obj).value
                    
                    end select ' select case vtag(0)
                
                end if ' if trim(ome.controls(obj).tag) <> "" then
            
            next obj ' for obj = 0 to totalobj
        
        end if ' if len(trim(validacamposobrigatorios)) > 0 then
        
        beditmode = false

end sub

public function validacamposobrigatorios(byref ome as object) as string
dim obj as integer
dim totalobj as integer
dim vtag as variant
dim btag as boolean
dim scamposembranco as string
dim stipocampo as string

    'application.volatile
    scamposembranco = ""

    totalobj = ome.controls.count - 1
    redim cformatfields(0 to totalobj)
    
    ' valida campos obrigatórios
    for obj = 0 to totalobj
    
        if trim(ome.controls(obj).tag) <> "" then
            
            if trim(application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 7, 0)) = 1 then
            
                stipocampo = trim(application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 6, 0))
                
                ' retorna a tag de metadados do campo na tabela de configurações
                if ome.controls(obj).value = "" then
                    
                   ome.controls(obj).backcolor = &h8080ff
                    vtag = split(ome.controls(obj).tag, "|")
                    scamposembranco = scamposembranco & chr(10) & chr(13) & vtag(1)
                                  
                elseif stipocampo <> "checklist" then
                
                    ome.controls(obj).backcolor = &h80000005
                    
                end if

            end if
            
        end if
    
    next obj ' for obj = 0 to totalobj
    
    validacamposobrigatorios = scamposembranco
end function

public sub initialize(byref ome as object)
dim obj as integer
dim totalobj as integer
dim vtag as variant
dim stag as string
    on error goto errhandler
    totalobj = ome.controls.count - 1
    redim cformatfields(0 to totalobj)
    
    for obj = 0 to totalobj
    
        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if trim(ome.controls(obj).tag) <> "" then
                
            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 13, 0)
            
            ' sepera as tags em uma lista
            vtag = split(stag, "|")

            select case vtag(0)

                case "date" ' formata a máscara de data automaticamente

                    set cformatfields(obj).formatdate = ome.controls(obj)
                    ome.controls(obj).maxlength = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 8, 0)

                case "time" ' formata a máscara de hora automaticamente

                    set cformatfields(obj).formattime = ome.controls(obj)
                    ome.controls(obj).maxlength = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 8, 0)
                    
                case "memo"
                    ome.controls(obj).maxlength = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 8, 0)
                    
                 case "senha"
                    ome.controls(obj).passwordchar = "*"
                    
                case "lista" ' carrega no combobox os ítens separados por ";" no campo "systablefields.valores da lista"
                
                    dim slistaitens as string
                    dim vlista as variant
                    dim i as integer
                    
                    vlista = split(application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 14, 0), ";")
                    
                    ome.controls(obj).clear
                    for i = lbound(vlista) to ubound(vlista)
                        
                        ome.controls(obj).additem trim(vlista(i))
                    
                    next i
                    
                case "combo", "checklist" ' carrega a lista de itens da tabela lookup na lista
                
                    dim rtablesource, rtablelookup as range
                    dim sfieldlookup as string
                    dim ifieldlist as integer
                    dim linha as integer
                    dim sstringaux as string

                    set rtablesource = range(vtag(4))
                    set rtablelookup = range(vtag(5))
                    sfieldlookup = vtag(6)
                    ifieldlist = application.worksheetfunction.match(sfieldlookup, rtablelookup, 0)

                    linha = 1
                    sstringaux = ""
                    ome.controls(obj).clear
                    do until rtablesource.cells(linha, ifieldlist) = ""
                        
                        if sstringaux <> rtablesource.cells(linha, ifieldlist) then
                            
                            ome.controls(obj).additem rtablesource.cells(linha, ifieldlist)
                            sstringaux = rtablesource.cells(linha, ifieldlist)
                        
                        end if
                        
                        linha = linha + 1
                    
                    loop
                    
                end select ' select case vtag(0)
            
            end if ' if trim(ome.controls(obj).tag) <> "" then
        
        next obj ' for obj = 0 to totalobj
        exit sub
errhandler:
        stop
        resume next
end sub

public sub carrega_treeview1(byval stablename as string, byval stablefields as string, byval ichave as integer, byval campostreeview as string, byref treeview as object)
dim vlistacampos, vcolunas() as variant
dim ilin as long
dim icampos, i, x, icolunaroot, icoluna, ilinha as integer
dim sfn_key, ssn_key, stn_key, stv_text, stv_key, stv_keyaux, scolunas, slinha as string
dim nrootnode, nchildnode as node

    vlistacampos = split(campostreeview, "|")
    redim preserve vcolunas(0 to ubound(vlistacampos)) as variant

    icampos = ubound(vlistacampos)
    ilin = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).row

    treeview.nodes.clear
    scolunas = ""

    for i = 0 to icampos

       if i = 0 then

            vcolunas(i) = application.worksheetfunction.match(vlistacampos(i), range(stablefields), 0)
            icolunaroot = application.worksheetfunction.match(vlistacampos(i), range(stablefields), 0)
            stv_keyaux = ""

            ' percorre todos os registros da tabela em busca dos valores únicos para o root node
            for x = 2 to ilin

                stv_key = worksheets(stablename).cells(x, icolunaroot).text

                if stv_key <> stv_keyaux then

                    sfn_key = "fn_" & worksheets(stablename).cells(x, vcolunas(i))
                    stv_text = worksheets(stablename).cells(x, vcolunas(i))

                    treeview.nodes.add key:=sfn_key, _
                                       text:=stv_text

                    stv_keyaux = cstr(worksheets(stablename).cells(x, vcolunas(i)))

                    end if ' if stv_key <> stv_keyaux then

                next x ' for i = 0 to ilin
            end if


        next i ' for i = 0 to icampos

    exit sub

    treeview.nodes.expanded = true


    'caregas a quantidade de linhas da tabela

end sub

public sub carrega_treeview(byval stablename as string, byval stablefields as string, byval ichave as integer, byval campostreeview as string, byref treeview as object)
dim vlistacampos, vcolunas() as variant
dim ilin as long
dim icampos, i, x, icolunaroot, icoluna, ilinha as integer
dim sfn_key, ssn_key, stn_key, stv_text, stv_key, stv_keyaux, scolunas, slinha as string
dim nrootnode, nchildnode as node

    vlistacampos = split(campostreeview, "|")
    redim preserve vcolunas(0 to ubound(vlistacampos)) as variant

    icampos = ubound(vlistacampos)
    ilin = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).row

    treeview.nodes.clear
    scolunas = ""

    for i = 0 to icampos

        select case i

            case 0

                vcolunas(i) = application.worksheetfunction.match(vlistacampos(i), range(stablefields), 0)
                icolunaroot = application.worksheetfunction.match(vlistacampos(i), range(stablefields), 0)
                stv_keyaux = ""

                ' percorre todos os registros da tabela em busca dos valores únicos para o root node
                for x = 2 to ilin

                    stv_key = worksheets(stablename).cells(x, icolunaroot).text

                    if stv_key <> stv_keyaux then

                        sfn_key = "fn_" & worksheets(stablename).cells(x, vcolunas(i))
                        stv_text = worksheets(stablename).cells(x, vcolunas(i))

                        treeview.nodes.add key:=sfn_key, _
                                           text:=stv_text

                        stv_keyaux = cstr(worksheets(stablename).cells(x, vcolunas(i)))

                        end if ' if stv_key <> stv_keyaux then

                    next x ' for i = 0 to ilin

            case 1

                vcolunas(i) = application.worksheetfunction.match(vlistacampos(i), range(stablefields), 0)
                stv_keyaux = ""

                for x = 2 to ilin

                    stv_key = worksheets(stablename).cells(x, vcolunas(i)).text
                    if stv_key <> stv_keyaux then

                        if i = ubound(vcolunas) then
                            slinha = cstr(worksheets(stablename).cells(x, 1)) & "|"
                        end if

                        ilinha = worksheets(stablename).cells(x, ichave)
                        sfn_key = "fn_" & worksheets(stablename).cells(x, vcolunas(i - 1))
                        ssn_key = slinha & "sn_" & worksheets(stablename).cells(x, vcolunas(i))
                        stv_text = worksheets(stablename).cells(x, vcolunas(i))

                        treeview.nodes.add sfn_key, _
                                           tvwchild, _
                                           key:=ssn_key, _
                                           text:=stv_text

                        stv_keyaux = worksheets(stablename).cells(x, vcolunas(i))

                        end if

                    next x ' for i = 0 to ilin

            case 2
                'exit sub

                vcolunas(i) = application.worksheetfunction.match(vlistacampos(i), range(stablefields), 0)
                stv_keyaux = ""

                for x = 2 to ilin

                    stv_key = worksheets(stablename).cells(x, vcolunas(i)).text
                    'if stv_key <> stv_keyaux then

                        if i = ubound(vcolunas) then
                            slinha = cstr(worksheets(stablename).cells(x, 1)) & "|"
                        end if

                        ilinha = worksheets(stablename).cells(x, ichave)
                        ssn_key = "sn_" & worksheets(stablename).cells(x, vcolunas(i - 1))
                        stn_key = slinha & "sn_" & worksheets(stablename).cells(x, vcolunas(i))
                        stv_text = worksheets(stablename).cells(x, vcolunas(i))

                        treeview.nodes.add ssn_key, _
                                           tvwchild, _
                                           key:=stn_key, _
                                           text:=stv_text

                        stv_keyaux = worksheets(stablename).cells(x, vcolunas(i))

                        'end if

                    next x ' for i = 0 to ilin

            end select

        next i ' for i = 0 to icampos

    exit sub

    treeview.nodes.expanded = true


    'caregas a quantidade de linhas da tabela

end sub

public sub savereg(byref ome as object, byval stablename as string, byval stablefields as string, byref otreeview as object, byval beditmode as boolean)
dim obj, totalobj, icoluna, ilinha as integer
dim llastrow, lnewregid as long
dim vtag, vtable, vlinha as variant
dim stag as string

    totalobj = ome.controls.count - 1

    if beditmode then
        
        vlinha = split(otreeview.selecteditem.key, "|")
        ilinha = vlinha(0)
    
    else
        llastrow = worksheets(stablename).cells(worksheets(stablename).cells.rows.count, 1).end(xlup).offset(0, 0).row
        lnewregid = llastrow + 1
        
        ilinha = lnewregid
        end if
                
    'worksheets(stablename).cells(ilinha, 1) = ilinha 'application.worksheetfunction.match(ilinha, range(stableid), 1)
        
    for obj = 0 to totalobj
    
        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if trim(ome.controls(obj).tag) <> "" then
            
            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 13, 0)
            
            ' sepera as tags em uma lista
            vtag = split(stag, "|")
            
            vtable = split(ome.controls(obj).tag, "|")
            icoluna = application.worksheetfunction.match(vtable(1), range(stablefields), 0)
            
            'range(stablename).rows.count + 1

            select case vtag(0)
            
                case "tab"
                
                    worksheets(stablename).cells(ilinha, icoluna) = ome.controls(obj).value
                    
                case "money"
                    dim cmoney as currency
                    
                    cmoney = ome.controls(obj).value
                    
                    worksheets(stablename).cells(ilinha, icoluna) = cmoney
                    
                case "calculado"
                
                    worksheets(stablename).cells(ilinha, icoluna) = ome.controls(obj).value
                
                case "value"
                    dim cvalue as double
                    
                    if trim(ome.controls(obj).value) <> "" then
                        cvalue = ome.controls(obj).value
                    
                        worksheets(stablename).cells(ilinha, icoluna) = cvalue
                    end if
                
                case "checklist" ' trata os campos checklist

                    dim spicklist, spicklist2 as string
                    dim i as integer

                    spicklist = ""
                    spicklist2 = ""
                    for i = 0 to ome.controls(obj).listcount - 1

                       if ome.controls(obj).selected(i) = true then

                          if spicklist = "" then
                             spicklist = ome.controls(obj).list(i)
                           else
                             spicklist2 = ome.controls(obj).list(i)
                             spicklist = spicklist & ";" & spicklist2
                           end if

                        end if

                    next

                    worksheets(stablename).cells(ilinha, icoluna) = spicklist
                    
                case else
                    
                    worksheets(stablename).cells(ilinha, icoluna) = ome.controls(obj).value
                
                end select ' select case vtag(0)
            
            end if ' if trim(me.controls(obj).tag) <> "" then
        
        next obj ' for obj = 0 to totalobj
    
end sub

public sub refreshpowequery()

    'activeworkbook.refreshall
    'activeworkbook.connections("thisworkbookdatamodel").refresh
    
end sub

public sub mytreeview_findnode(byref ome as object, byval strkey as string)
dim mynode as node

    for each mynode in ome.tvwativos.nodes
   
        mynode.bold = mynode.key = strkey
        
    next
end sub

public sub enabledisablecontrols(byref ome as object, byval benabled as boolean)
dim obj, totalobj as integer
dim stag as string
dim vtag as variant

const lenablecolor as long = &h80000005
const ldesablecolor as long = &h8000000f

    totalobj = ome.controls.count - 1
    redim cformatfields(0 to totalobj)

    for obj = 0 to totalobj

        ' verifica se a propriedade tag tem conteúdo e se o campo é físico
        if len(trim(ome.controls(obj).tag)) > 0 then

            ' retorna a tag de metadados do campo na tabela de configurações
            stag = application.worksheetfunction.vlookup(ome.controls(obj).tag, range("systablefields"), 13, 0)

            ' sepera as tags em uma lista
            vtag = split(stag, "|")

             select case vtag(0)

                case "tab", "boolean"

                    if benabled then

                        ome.controls(obj).enabled = benabled
                        ome.controls(obj).backcolor = &h8000000f

                    else

                        ome.controls(obj).enabled = false

                    end if

                case else

                    if benabled then

                        ome.controls(obj).enabled = benabled
                        ome.controls(obj).backcolor = lenablecolor

                    else

                        ome.controls(obj).enabled = false
                        ome.controls(obj).backcolor = ldesablecolor

                    end if

            end select

        end if

    next obj ' for obj = 0 to totalobj
end sub

public sub sortmultiplecolumns(byval stablename as string, byval scampos as string)
dim ws as worksheet
dim tbl as listobject
dim vlistacampos as variant
dim icampos, i as integer

    set ws = sheets(stablename)
    set tbl = ws.listobjects(stablename)

    vlistacampos = split(scampos, "|")
    icampos = ubound(vlistacampos)
    
    with tbl.sort
        
        .sortfields.clear
        
        for i = 0 to icampos

            .sortfields.add key:=range(stablename & "[" & vlistacampos(i) & "]"), sorton:=xlsortonvalues, order:=xlascending
        
        next i
        
        .header = xlyes
        .apply
        
    end with
    
end sub


