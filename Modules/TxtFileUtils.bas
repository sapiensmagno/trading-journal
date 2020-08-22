Attribute VB_Name = "TxtFileUtils"
 Option Explicit
 Option Compare Text
 
 Sub StringToFile(ByVal str As String, Path As String, Optional Mode As String = "Overwrite")
    
    Dim FileNum As Integer
    FileNum = FreeFile ' next free filenumber
    'creates the new file
    select case mode
        case "overwrite"
            open path for output as #filenum
        case "append"
            open path for append as #filenum
    end select
    print #filenum, str
    close #filenum ' close the file
    
end sub

sub copyfolder(byval src as string, byval dest as string, optional byval overwrite as boolean)
    dim fso as object
    set fso = createobject("scripting.filesystemobject")
    fso.copyfolder src, dest, overwrite
    set fso = nothing
end sub

sub copyfile(byval src as string, byval dest as string, optional byval overwrite as boolean)
    dim fso as object
    set fso = createobject("scripting.filesystemobject")
    fso.copyfile source:=src, destination:=dest, overwritefiles:=overwrite
    set fso = nothing
end sub

public function comparetextfiles(file1path as string, file2path as string) as boolean
'''''''''''''''''
'http://quadexcel.com/vba-macro-to-compare-two-files-to-determine-if-they-are-identical/
'''''''''''''''''''
'**********************************************************
'‘purpose: check to see if two files are identical
'‘file1 and file2 = fullpaths of files to compare
'‘will compare complete content of the file, including length of the document (bit to bit)
'‘**********************************************************

dim file1content, file2content as variant

file1content = getfilecontent(file1path)
file2content = getfilecontent(file2path)

comparetextfiles = file1content = file2content

end function

function getfilecontent(name as variant) as variant
dim intunit as integer

on error goto errgetfilecontent
intunit = freefile
open name for input as intunit
getfilecontent = input(lof(intunit), intunit)

errgetfilecontent:
close intunit
exit function
end function

sub deletefile(byval filetodelete as string)
   if fileexists(filetodelete) then
      setattr filetodelete, vbnormal
      kill filetodelete
   end if
end sub

function fileexists(byval filetotest as string) as boolean
   fileexists = (dir(filetotest) <> "")
end function

function folderexists(byval folderpath as string) as boolean
    dim fso
    set fso = createobject("scripting.filesystemobject")
    folderexists = fso.folderexists(folderpath)
end function


