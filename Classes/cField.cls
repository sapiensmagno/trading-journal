VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

implements ifield

'@classe
private sub class_initialize()
    err.raise 5 'todo implement interface member
end sub

private sub class_terminate()
    err.raise 5 'todo implement interface member
end sub

'@interface
private property get ifield_id() as long
    err.raise 5 'todo implement interface member
end property

private property let ifield_table(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_table() as string
    err.raise 5 'todo implement interface member
end property

private property get ifield_name() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_name(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_description() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_description(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_datatype() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_datatype(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_size() as integer
    err.raise 5 'todo implement interface member
end property

private property let ifield_size(byval rhs as integer)
    err.raise 5 'todo implement interface member
end property

private property get ifield_mask() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_mask(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_tablesource() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_tablesource(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_listoffields() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_listoffields(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_fieldlookup() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_fieldlookup(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_tag() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_tag(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_valuesoflist() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_valuesoflist(byval rhs as string)
    err.raise 5 'todo implement interface member
end property

private property get ifield_treeviewfieldorder() as integer
    err.raise 5 'todo implement interface member
end property

private property let ifield_treeviewfieldorder(byval rhs as integer)
    err.raise 5 'todo implement interface member
end property

private property get ifield_treeviewfieldlist() as string
    err.raise 5 'todo implement interface member
end property

private property let ifield_treeviewfieldlist(byval rhs as string)
    err.raise 5 'todo implement interface member
end property



