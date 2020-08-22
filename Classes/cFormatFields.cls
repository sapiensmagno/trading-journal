VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFormatFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

public withevents formatdate as msforms.textbox
attribute formatdate.vb_varhelpid = -1
public withevents formattime as msforms.textbox
attribute formattime.vb_varhelpid = -1
public withevents formatvalue as msforms.textbox
attribute formatvalue.vb_varhelpid = -1
public withevents formatmoney as msforms.textbox
attribute formatmoney.vb_varhelpid = -1

private sub formatdate_keypress(byval keyascii as msforms.returninteger)
dim strvalid as string

    strvalid = "0123456789"
    
    if instr(strvalid, chr(keyascii)) = 0 then
        keyascii = 0
    else
        'formatdate.maxlength = 10
        
        if formatdate.selstart = 2 then
            formatdate.seltext = "/"
        elseif formatdate.selstart = 5 then
            formatdate.seltext = "/"
        end if
    end if

end sub

private sub formattime_keypress(byval keyascii as msforms.returninteger)
dim strvalid as string

    strvalid = "0123456789"
    
    if instr(strvalid, chr(keyascii)) = 0 then
        keyascii = 0
    else
        'formattime.maxlength = 8
        
        if formattime.selstart = 2 then
            formattime.seltext = ":"
        elseif formattime.selstart = 5 then
            formattime.seltext = ":"
        elseif formattime.selstart = 8 then
            formattime.seltext = ":"
        end if
    end if
    
end sub

