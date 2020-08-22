VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

implements itable

'@classe
private sub class_initialize()
    err.raise 5 'todo implement interface member
end sub

private sub class_terminate()
    err.raise 5 'todo implement interface member
end sub

'@interface
private property get itable_id() as long
    err.raise 5 'todo implement interface member
end property

private property let itable_listoffields(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get itable_listoffields() as string
    err.raise 5 'todo implement interface member
end property

private property get itable_name() as string
    err.raise 5 'todo implement interface member
end property

private property let itable_name(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get itable_tabletype() as string
    err.raise 5 'todo implement interface member
end property

private property let itable_tabletype(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

