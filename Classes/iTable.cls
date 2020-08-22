VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

private type ttable
    id as long
    tabletype as string
    name as string
    listoffields as string
end type

private this as ttable

public property get id() as long
    id = this.id
end property

public property get tabletype() as string
    tabletype = this.tabletype
end property

public property let tabletype(byval value as string)
    this.tabletype = value
end property

public property get name() as string
    name = this.name
end property

public property let name(byval value as string)
    this.name = value
end property

public property get listoffields() as string
    listoffields = this.listoffields
end property

public property let listoffields(byval value as string)
    this.listoffields = value
end property

