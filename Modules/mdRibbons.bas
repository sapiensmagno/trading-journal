Attribute VB_Name = "mdRibbons"
Option Explicit

Sub btnAdicionarTrade_Click(control As IRibbonControl)
    
    frmTradesCadastro.Show 0
    
End Sub

sub btneditartrade_click(control as iribboncontrol)

    with frmtradescadastro
        
        .caption = "revisão/edição de trade"
        .lbltitle.caption = "revisão/edição de trade"
        .cbxtrade.visible = true
        .cbxtrade.left = .txtnrotrade.left
        .txtnrotrade.visible = false
        .mpgnewtrade.pages(3).visible = true
        .show
        
    end with
    
end sub

sub btncorretorascadastro_click(control as iribboncontrol)

    frmcorretorascadastro.show
    
end sub

sub btnativoscadastro_click(control as iribboncontrol)

    frmativoscadastro.show
    
end sub

sub btnativosconfiguracao_click(control as iribboncontrol)

    frmativosconfiguracao.show
    
end sub

sub btnestrategiascadastro_click(control as iribboncontrol)

    frmestrategiascadastro.show
    
end sub

sub btnconfiguracaogeral_click(control as iribboncontrol)

    frmconfiguracoes.show

end sub

sub btnconfirurarbase_click(control as iribboncontrol)
    
    frmsysbuilder.show
    
end sub

sub btnchecklistconfig_click(control as iribboncontrol)
    
    frmchecklistcadastro.show
    
end sub

sub btnchecklistdiario_click(control as iribboncontrol)
    
    frmchecklistdiario.show 0
    
end sub


sub btnimportartrades_click(control as iribboncontrol)
    
    frmimportartrades.show
    
end sub

sub btnplanodetrade_click(control as iribboncontrol)
    
    frmplanodetrade.show
    
end sub

sub btndeposito_click(control as iribboncontrol)
    
    frmcashflow.show
    
end sub

sub btnshowdashboard_click(control as iribboncontrol)
    
    sheets("dashboard").select
    
end sub

sub btnrpt_evolucao_click(control as iribboncontrol)
    
    sheets("evolução").select
    
end sub

sub btnrpt_desempenho_click(control as iribboncontrol)
    
    sheets("métricas de desempenho").select
    
end sub

sub btnrpt_lucroprejuizo_click(control as iribboncontrol)
    
    sheets("métricas de lp").select
    
end sub

sub btnrpt_volume_click(control as iribboncontrol)
    
    sheets("métricas de volume").select
    
end sub

sub btnrpt_custos_click(control as iribboncontrol)
    
    sheets("métricas de custos").select
    
end sub

sub btnrpt_pivottable_click(control as iribboncontrol)
    
    sheets("análise personalizada").select
    
end sub







