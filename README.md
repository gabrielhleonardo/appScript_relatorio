# appScript_relatorio
PROJETO : Planilha elaborada com App Script para cadastro de itens e operações, gerar e enviar relatório por email. Código Abaixo:
=============================================================================================
var planilha = SpreadsheetApp.getActiveSpreadsheet();

var cadastro = planilha.getSheetByName("Cadastro");
var baseDados = planilha.getSheetByName("BaseDados"); 
var movimentacoes = planilha.getSheetByName("Movimentações");
var gerador = planilha.getSheetByName("Gerador de relatórios");
var relatorio = planilha.getSheetByName("Relatório");

function cadastrar() {

  var data = cadastro.getRange("C3:G3").getValue();
  var tipo = cadastro.getRange("C5").getValue();
  var categoria = cadastro.getRange("F5:G5").getValue();
  var descricao = cadastro.getRange("C7:G7").getValue();
  var valor = cadastro.getRange("C9:G9").getValue();
  var ultimaLinha = baseDados.getLastRow()+1;

  baseDados.getRange(ultimaLinha,1).setValue(data);
  baseDados.getRange(ultimaLinha,2).setFormula('=SPLIT(A'+ultimaLinha+';"/")');
  baseDados.getRange(ultimaLinha,5).setValue(tipo);
  baseDados.getRange(ultimaLinha,6).setValue(categoria);
  baseDados.getRange(ultimaLinha,7).setValue(descricao);
  
  if(tipo == "Entrada"){
     baseDados.getRange(ultimaLinha,8).setValue(valor); 
  } else {
    baseDados.getRange(ultimaLinha,8).setValue(-valor);  
  }

  if(ultimaLinha == 2){
      movimentacoes.getRange(ultimaLinha,9).setFormula("=H2");
  } else {
        movimentacoes.getRange(ultimaLinha,9).setFormula("I"+(ultimaLinha-1)+ "+H"+ultimaLinha+"");
  }
  
  limpar();
}

function limpar() {

  cadastro.getRange("C3:G3").clearContent();//data
  cadastro.getRange("C5").clearContent();//tipo
  cadastro.getRange("F5:G5").clearContent();//categoria
  cadastro.getRange("C7:G7").clearContent();//descricao
  cadastro.getRange("C9:G9").clearContent();//valor
}


function gerar() {
  relatorio.getRange("F2:F").clearContent();
  relatorio.getRange("F2").setFormula("=E2");

  for(var i = 3; i <=relatorio.getLastRow(); i++) {
    relatorio.getRange(i,6).setFormula("=F"+(i-1)+ "+E"+i+"");
  }
}


function enviar(){

var destinatario = gerador.getRange("K4:K5").getValue();
var mensagem = gerador.getRange("I4:I6").getValue();

var email = {
  to: destinatario,
  subject:  "Relatório Financeiro",
  body: mensagem,
  name: "Nome da Planilha",
  attachments: [planilha.getAs(MimeType.PDF).setName("Nome do Relatorio.pdf")]
}

cadastro.hideSheet();
movimentacoes.hideSheet();
gerador.hideSheet();

MailApp.sendEmail(email);

gerador.getRange("K4:K5").clearContent();
gerador.getRange("I4:I6").clearContent();
gerador.getRange("C3").clearContent();
gerador.getRange("F3").clearContent();
gerador.getRange("C5:F5").clearContent();

cadastro.showSheet();
movimentacoes.showSheet();
gerador.showSheet();

}
