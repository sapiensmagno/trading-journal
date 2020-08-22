VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCorretoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit


private type tcorretora
    id as long
    nacionalinternacional  as string
    nome as string
    fone01 as string
    fone02 as string
    whatsapptelegram as string
    website as string
    contanro as string
    senha as string
    assinaturaeletronica as string
    emailcadastro as string
    portalhb as string
    temacessor as boolean
    assessornome as string
    assessorcodigo as string
    assessorcontato as string
    assessorfone01 as string
    assessorfone02 as string
    assessorwhatsapptelegram as string
    assessoremail as string
    ações as boolean
    futuros as boolean
    opcoes as boolean
    termo as boolean
    fii as boolean
    etf as boolean
    commodities as boolean
    dadosdeposito as string
    deleted as boolean
    editado as date
end type

private this as tcorretora

public property get id() as long
    
    id = this.id
    
end property

public property get nacionalinternacional() as string
    
    nacionalinternacional = this.nacionalinternacional
    
end property

public property let nacionalinternacional(byval value as string)
    
    this.nacionalinternacional = value
    
end property

public property get nome() as string
    
    nome = this.nome
    
end property

public property let nome(byval value as string)
    
    this.nome = value
    
end property




