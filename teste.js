function onEdit(e) {
  var sheetOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Strikes e Suspensões");
  var sheetDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("block permanente backoffice");

  var colunaEntregador = 1; // Coluna "Entregador" (A)
  var colunaTelefone = 2;   // Coluna "Telefone" (B)
  var colunaEntrega = 3;    // Coluna "Entrega/pedido" (C)
  var colunaData = 4;       // Coluna "Data Ocorrência" (D)
  var colunaAcao = 5;       // Coluna "AÇÃO" (E)

  var linhaEditada = e.range.getRow();
  var colunaEditada = e.range.getColumn();

  // Verifica se a edição foi na coluna "AÇÃO"
  if (colunaEditada !== colunaAcao) return;

  var valorAcao = e.range.getValue();
  var valoresOrigem = [
    sheetOrigem.getRange(linhaEditada, colunaEntregador).getValue(), // Entregador
    sheetOrigem.getRange(linhaEditada, colunaTelefone).getValue(),   // Telefone
    sheetOrigem.getRange(linhaEditada, colunaEntrega).getValue(), // Entrega      
    sheetOrigem.getRange(linhaEditada, colunaData).getValue()// Data Ocorrência
  ]

  if (valorAcao === "BLOQUEIO PERMANENTE") {
    atualizarOuAdicionarLinha(sheetDestino, valoresOrigem);
  } 
  if (valorAcao === "strike") {
    atualizarDataOcorrencia(sheetOrigem, linhaEditada, valoresOrigem);
  }
  if (valorAcao != "strike") {
    removerDataOcorrencia(sheetOrigem, linhaEditada);
  }
  else {
    removerLinhaCorrespondente(sheetDestino, valoresOrigem);
  }
}

//Atualiza data ocorrência
function atualizarDataOcorrencia(sheetOrigem, linhaEditada, valoresOrigem){
    sheetOrigem.getRange(linhaEditada, 9).setValue(valoresOrigem[3]);
  }

//Remove data ocorrência
function removerDataOcorrencia(sheetOrigem, linhaEditada) {
    sheetOrigem.getRange(linhaEditada, 9).setValue("");
}
// Atualiza ou adiciona uma nova linha com os dados na planilha de destino
function atualizarOuAdicionarLinha(sheetDestino, valoresOrigem) {
  var linhasDestino = sheetDestino.getDataRange().getValues();
  var linhaExistente = linhasDestino.findIndex(linha => 
    linha[0] == valoresOrigem[0] && linha[1] == valoresOrigem[1] && linha[2] == valoresOrigem[2]);

  if (linhaExistente !== -1) {
    // Atualiza linha existente
    sheetDestino.getRange(linhaExistente + 1, 1, 1, 5).setValues([[...valoresOrigem, "NÃO"]]);
  } else {
    // Adiciona nova linha no final
    var ultimaLinhaDestino = sheetDestino.getLastRow() + 1;
    sheetDestino.getRange(ultimaLinhaDestino, 1, 1, 5).setValues([[...valoresOrigem, "NÃO"]]);
  }
}

// Remove uma linha correspondente, se encontrada
function removerLinhaCorrespondente(sheetDestino, valoresOrigem) {
  var linhasDestino = sheetDestino.getDataRange().getValues();

  for (var i = 0; i < linhasDestino.length; i++) {
    if (linhasDestino[i][0] == valoresOrigem[0] && linhasDestino[i][1] == valoresOrigem[1] && linhasDestino[i][2] == valoresOrigem[2]) {
      sheetDestino.deleteRow(i + 1); // Ajuste do índice (base 1)
      break;
    }
  }
}
