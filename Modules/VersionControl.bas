Attribute VB_Name = "VersionControl"
Option Explicit
Option Compare Text

' With adaptations from: https://github.com/jonathanng/art_vandelay
Private Const DOCUMENT_FOLDER = "Archive\"
Private Const VBACODE_FOLDER = "code\"
Private Const TEMP_ZIP = "\temp.zip"
Private Const EXTENSION = "xlsm"
const vbext_ct_classmodule = 2
const vbext_ct_document = 100
const vbext_ct_msform = 3
const vbext_ct_stdmodule = 1

private function importfromfolder() as string

    importfromfolder = "c:\users\virtu\dropbox\trading\mentoring cadu\my trading journal\"

end function

private function exporttofolder() as string

    exporttofolder = "c:\users\virtu\dropbox\trading\mentoring cadu\my trading journal\" 'thisworkbook.path & "\"

end function

public sub cleanup()
    dim vbcomp as vbcomponent
    for each vbcomp in activeworkbook.vbproject.vbcomponents
        if right(vbcomp.name, 1) = "1" then vbcomp.name = left(vbcomp.name, len(vbcomp.name) - 1)
    next
    debug.print "cleanup done!"

end sub

public sub export()

    'http://www.pretentiousname.com/excel_extractvba/index.html

    dim vbcomp as vbcomponent
    dim path as string

    'ensuredir folder
    'ensuredir folder & vbacode_folder

    for each vbcomp in activeworkbook.vbproject.vbcomponents

        if ext(vbcomp) <> "" then
'            ensuredir folder & vbacode_folder & subfolder(vbcomp)
            path = exporttofolder & vbacode_folder & subfolder(vbcomp) & "\" & vbcomp.name & ext(vbcomp)
            debug.print "exporting " & path
            vbcomp.export path
        end if

    next
    codetolowercase
    debug.print "exporting done!"

end sub

public sub import()

    dim vbcomp as vbcomponent
    dim path as string
    dim file as long
    dim line as long
    dim txt_line as string
    dim code as string
    dim fso as filesystemobject
    dim fso_folder as variant
    dim fso_subfolder as variant
    dim fso_file as variant
    
    set fso = createobject("scripting.filesystemobject")
    set fso_folder = fso.getfolder(importfromfolder & vbacode_folder)
    
    removeall
    end
    'para evitar erros: 1) rode o método import até o end. comente o end. rode novamente.
    
    'se estiver em /documents verificar se o componente ja existe na pasta de trabalho
    '   se nao existe, importa
    '   se ja existe, apaga as linhas do componente, le o arquivo e escreve no componente a partir da 10a linha
    'para as outras pastas, importa o modulo (tudo ja terá sido previamente removido)
    
    ' insert breakpoint here, stop execution and then run the code again.
    ' this for each often starts before "removeall" has finished and causes bugs.
    for each fso_subfolder in fso_folder.subfolders
        select case fso_subfolder.name
            case "documents"
            ' document modules (sheetx and thisworkbook) cannot be removed. so, delete all code in that
            ' component and add the lines from the source file back in to the module.                                                                      '
                for each fso_file in fso_subfolder.files
                    debug.print "write to document "; fso_file.path
                    on error goto errhandler
                    set vbcomp = thisworkbook.vbproject.vbcomponents(left(fso_file.name, len(fso_file.name) - 4))
                    if not vbcomp is nothing then
                        file = freefile
                        open fso_file.path for input as #file
                        do until eof(file)
                            line = line + 1
                            line input #file, txt_line
                            if line > 9 then code = code & txt_line & vbcrlf 'code begins at line 10. before it, only vb attributes not visible inside the vba editor.
                        loop
                        with vbcomp.codemodule
                            .deletelines 1, .countoflines
                            .insertlines 1, code
                        end with
                        line = 0
                        code = ""
                        txt_line = ""
                        close #file
                    end if
                next fso_file
            
            case else
                for each fso_file in fso_subfolder.files
                    if (right(fso_file.name, 3) = "bas" or _
                        right(fso_file.name, 3) = "cls" or _
                        right(fso_file.name, 3) = "frm") then 'everthing should be ignored, except cls, bas frm
                        
                        debug.print "importing "; fso_file.path
                        thisworkbook.vbproject.vbcomponents.import fso_file.path
                        
                    end if
                next fso_file
            
        end select
        
    next fso_subfolder

    debug.print "importing done!"
    
    exit sub
     
errhandler:
    if err.number = 9 then
        debug.print fso_file.name; " doesnt exist. "
        set vbcomp = nothing
        resume next
    else
        debug.print fso_file.name; " - unexpected error. stopping. "; err.description
        resume next
    end if
    
end sub

public sub removeall()
    
    dim vbcomp as vbcomponent
    
    for each vbcomp in thisworkbook.vbproject.vbcomponents
        if vbcomp.name <> "versioncontrol" and vbcomp.name <> "ufobjbrowser" and vbcomp.type <> vbext_ct_document then
            thisworkbook.vbproject.vbcomponents.remove vbcomp
        end if
    next
    
    debug.print "modules removed. "

end sub
private function ext(vbcomp as vbcomponent) as string

    select case vbcomp.type
        case vbext_ct_classmodule:     ext = ".cls"
        case vbext_ct_document:        ext = ".cls"
        case vbext_ct_msform:          ext = ".frm"
        case vbext_ct_stdmodule:       ext = ".bas"
        case else:                     ext = ""
    end select

end function

private function subfolder(vbcomp as vbcomponent) as string

    select case vbcomp.type
         case vbext_ct_classmodule:     subfolder = "classes"
         case vbext_ct_document:        subfolder = "documents"
         case vbext_ct_msform:          subfolder = "forms"
         case vbext_ct_stdmodule:       subfolder = "modules"
         case else:                     subfolder = ""
    end select

end function

public sub codetolowercase()

    dim vbcomp as vbcomponent
    dim path as string
    dim file as long
    dim line as long
    dim txt_line as string
    dim code as string
    dim fso as filesystemobject
    dim fso_folder as variant
    dim fso_subfolder as variant
    dim fso_file as variant
        
    set fso = createobject("scripting.filesystemobject")
    set fso_folder = fso.getfolder(importfromfolder & vbacode_folder)
    
    for each fso_subfolder in fso_folder.subfolders

        for each fso_file in fso_subfolder.files
            on error goto errhandler
            file = freefile
            open fso_file.path for input as #file
            do until eof(file)
                line = line + 1
                line input #file, txt_line
                if line > 9 then
                    code = code & lcase(txt_line) & vbcrlf 'code begins at line 10. before it, only vb attributes not visible inside the vba editor.
                else
                    code = code & txt_line & vbcrlf
                end if
            loop
            close #file
            txtfileutils.stringtofile code, fso_file.path
            line = 0
            code = ""
            txt_line = ""
        next fso_file
    next fso_subfolder
    
    exit sub
     
errhandler:
    if err.number = 9 then
        debug.print fso_file.name; " doesnt exist. "
        set vbcomp = nothing
        resume next
    else
        debug.print fso_file.name; " - unexpected error. stopping. "; err.description
        resume next
    end if
    
end sub




