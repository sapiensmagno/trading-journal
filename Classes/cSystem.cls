VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

private sub carrega_treeview(byval tabela as string, byval chave as string, byval campostreeview as string, byref treeview as object)
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
            
                        stv_keyaux = worksheets(stablename).cells(x, vcolunas(i))
                        
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
                        
                        ilinha = worksheets(stablename).cells(x, 1)
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
                        
                        ilinha = worksheets(stablename).cells(x, 1)
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


