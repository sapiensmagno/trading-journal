VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTrades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
option explicit

'implements itrades

private type ttrades
    id as long
    dataentrada as date
    horaentrada as date
    corretora as string
    ativo as string
    estrategia as string
    compravenda as string
    ehdaytrade as boolean
    precoentrada as double
    nrocontratos as integer
    precosl as double
    precotp as double
    motivacaoentrada as string
    datasaida as date
    horasaida as date
    precosaida as double
    lucroprejuizo as double
    custocorretagem as double
    custoemolumentos as double
    margem as double
    pontos as double
    motivacaosaida as string
    mep as double
    men as double
    tipooperacao as string
    tipoconta as string
    fezprecomedio as boolean
    revisartrade as boolean
    trademental as boolean
    condicaomercado as string
    tipooportunidade as string
    riscosempotencial as string
    riscoentrada as string
    riscosaida as string
    gestaodotrade as string
    imagem01 as string 'path
    imagem02 as string 'path
    observacoes as string
    traderevisado as boolean
    observacoesrevisao as string
    
    ' propriedades de sistema
    deleted as boolean
    'private editado as date
end type

private this as ttrades

public property get id() as long
    
    id = this.id
    
end property

public property get dataentrada() as date
    
    dataentrada = this.dataentrada
    
end property

public property let dataentrada(byval value as date)
    
    this.dataentrada = value
    
end property


public property get horaentrada() as date
    
    dataentrada = this.dataentrada
    
end property

public property let horaentrada(byval value as date)
    
    this.horaentrada = value
    
end property

public property get corretora() as string
    
    dataentrada = this.dataentrada
    
end property

public property let corretora(byval value as string)
    
    this.corretora = value
    
end property

public property get ativo() as string
    
    ativo = this.ativo
    
end property

public property let ativo(byval value as string)
    
    this.ativo = value
    
end property

public property get estrategia() as string
    
    estrategia = this.estrategia
    
end property

public property let estrategia(byval value as string)
    
    this.estrategia = value
    
end property

public property get compravenda() as string
    
    compravenda = this.compravenda
    
end property

public property let compravenda(byval value as string)
    
    this.compravenda = value
    
end property

public property get ehdaytrade() as boolean
    
    ehdaytrade = this.ehdaytrade
    
end property

public property let ehdaytrade(byval value as boolean)
    
    this.ehdaytrade = value
    
end property

public property get precoentrada() as double
    
    precoentrada = this.precoentrada
    
end property

public property let precoentrada(byval value as double)
    
    this.precoentrada = value
    
end property

public property get nrocontratos() as integer
    
    nrocontratos = this.nrocontratos
    
end property

public property let nrocontratos(byval value as integer)
    
    this.nrocontratos = value
    
end property

public property get precosl() as double
    
    precosl = this.precosl
    
end property

public property let precosl(byval value as double)
    
    this.precosl = value
    
end property

public property get precotp() as double
    
    precotp = this.precotp
    
end property

public property let precotp(byval value as double)
    
    this.precotp = value
    
end property

public property get motivacaoentrada() as string
    
    motivacaoentrada = this.motivacaoentrada
    
end property

public property let motivacaoentrada(byval value as string)
    
    this.motivacaoentrada = value
    
end property

public property get horasaida() as date
    
    horasaida = this.horasaida
    
end property

public property let horasaida(byval value as date)
    
    this.horasaida = value
    
end property

public property get precosaida() as double
    
    precosaida = this.precosaida
    
end property

public property let precosaida(byval value as double)
    
    this.precosaida = value
    
end property

public property get lucroprejuizo() as double
    
    lucroprejuizo = this.lucroprejuizo
    
end property

public property let lucroprejuizo(byval value as double)
    
    this.lucroprejuizo = value
    
end property

public property get custocorretagem() as double
    
    custocorretagem = this.custocorretagem
    
end property

public property let custocorretagem(byval value as double)
    
    this.custocorretagem = value
    
end property

public property get custoemolumentos() as double
    
    custoemolumentos = this.custoemolumentos
    
end property

public property let custoemolumentos(byval value as double)
    
    this.custoemolumentos = value
    
end property

public property get margem() as double
    
    margem = this.margem
    
end property

public property let margem(byval value as double)
    
    this.margem = value
    
end property

public property get pontos() as double
    
    pontos = this.pontos
    
end property

public property let pontos(byval value as double)
    
    this.pontos = value
    
end property

public property get motivacaosaida() as string
    
    motivacaosaida = this.motivacaosaida
    
end property

public property let motivacaosaida(byval value as string)
    
    this.motivacaosaida = value
    
end property

public property get mep() as double
    
    mep = this.mep
    
end property

public property let mep(byval value as double)
    
    this.mep = value
    
end property

public property get men() as double
    
    men = this.men
    
end property

public property let men(byval value as double)
    
    this.men = value
    
end property

public property get tipooperacao() as string
    
    tipooperacao = this.tipooperacao
    
end property

public property let tipooperacao(byval value as string)
    
    this.tipooperacao = value
    
end property

public property get tipoconta() as string
    
    tipoconta = this.tipoconta
    
end property

public property let tipoconta(byval value as string)
    
    this.tipoconta = value
    
end property

public property get fezprecomedio() as boolean
    
    fezprecomedio = this.fezprecomedio
    
end property

public property let fezprecomedio(byval value as boolean)
    
    this.fezprecomedio = value
    
end property

public property get revisartrade() as boolean
    
    revisartrade = this.revisartrade
    
end property

public property let revisartrade(byval value as boolean)
    
    this.revisartrade = value
    
end property

public property get trademental() as boolean
    
    trademental = this.trademental
    
end property

public property let trademental(byval value as boolean)
    
    this.trademental = value
    
end property

public property get condicaomercado() as string
    
    condicaomercado = this.condicaomercado
    
end property

public property let condicaomercado(byval value as string)
    
    this.condicaomercado = value
    
end property

public property get tipooportunidade() as string
    
    tipooportunidade = this.tipooportunidade
    
end property

public property let tipooportunidade(byval value as string)
    
    this.tipooportunidade = value
    
end property

public property get riscosempotencial() as string
    
    riscosempotencial = this.riscosempotencial
    
end property

public property let riscosempotencial(byval value as string)
    
    this.riscosempotencial = value
    
end property

public property get riscoentrada() as string
    
    riscoentrada = this.riscoentrada
    
end property

public property let riscoentrada(byval value as string)
    
    this.riscoentrada = value
    
end property

public property get riscosaida() as string
    
    riscosaida = this.riscosaida
    
end property

public property let riscosaida(byval value as string)
    
    this.riscosaida = value
    
end property

public property get gestaodotrade() as string
    
    gestaodotrade = this.gestaodotrade
    
end property

public property let gestaodotrade(byval value as string)
    
    this.gestaodotrade = value
    
end property

public property get imagem01() as string
    
    imagem01 = this.imagem01
    
end property

public property let imagem01(byval value as string)
    
    this.imagem01 = value
    
end property

public property get imagem02() as string
    
    imagem02 = this.imagem02
    
end property

public property let imagem02(byval value as string)
    
    this.imagem02 = value
    
end property

public property get observacoes() as string
    
    observacoes = this.observacoes
    
end property

public property let observacoes(byval value as string)
    
    this.observacoes = value
    
end property

'public property get trademental() as boolean
'
'    trademental = this.trademental
'
'end property
'
'public property let trademental(byval value as boolean)
'
'    this.trademental = value
'
'end property

public property get traderevisado() as boolean
    
    traderevisado = this.traderevisado
    
end property

public property let traderevisado(byval value as boolean)
    
    this.traderevisado = value
    
end property

public property let observacoesrevisao(byval value as string)
    
    this.observacoesrevisao = value
    
end property

public property get observacoesrevisao() as string
    
    observacoesrevisao = this.observacoesrevisao
    
end property

public property get deleted() as boolean
    
    deleted = this.deleted
    
end property

public property let deleted(byval value as boolean)
    
    this.deleted = value
    
end property

