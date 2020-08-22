VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

private type tfield
    id as long
    table as string
    name as string
    description as string
    datatype as string
    size as integer
    mask as string
    tablesource as string
    listoffields as string
    fieldlookup as string
    tag as string
    valuesoflist as string
    treeviewfieldorder as integer
    treeviewfieldlist as string
end type

private this as tfield

public property get id() as long
    id = this.id
end property

public property get table() as string
    table = this.table
end property

public property let table(byval value as string)
    this.table = value
end property

public property get name() as string
    name = this.name
end property

public property let name(byval value as string)
    this.name = value
end property

public property get description() as string
    description = this.description
end property

public property let description(byval value as string)
    this.description = value
end property

public property get datatype() as string
    datatype = this.datatype
end property

public property let datatype(byval value as string)
    this.datatype = value
end property

public property get size() as integer
    size = this.size
end property

public property let size(byval value as integer)
    this.size = value
end property

public property get mask() as string
    mask = this.mask
end property

public property let mask(byval value as string)
    this.mask = value
end property

public property get tablesource() as string
    tablesource = this.tablesource
end property

public property let tablesource(byval value as string)
    this.tablesource = value
end property

public property get listoffields() as string
    listoffields = this.listoffields
end property

public property let listoffields(byval value as string)
    this.listoffields = value
end property

public property get fieldlookup() as string
    fieldlookup = this.fieldlookup
end property

public property let fieldlookup(byval value as string)
    this.fieldlookup = value
end property

public property get tag() as string
    tag = this.tag
end property

public property let tag(byval value as string)
    this.tag = value
end property

public property get valuesoflist() as string
    valuesoflist = this.valuesoflist
end property

public property let valuesoflist(byval value as string)
    this.valuesoflist = value
end property

public property get treeviewfieldorder() as integer
    treeviewfieldorder = this.treeviewfieldorder
end property

public property let treeviewfieldorder(byval value as integer)
    this.treeviewfieldorder = value
end property

public property get treeviewfieldlist() as string
    treeviewfieldlist = this.treeviewfieldlist
end property

public property let treeviewfieldlist(byval value as string)
    this.treeviewfieldlist = value
end property


