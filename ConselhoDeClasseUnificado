function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Ferramentas')
      .addItem('Consolidar Bimestres', 'ConsolidaBimestres')
      .addToUi();
}

function ConsolidaBimestres() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const paginaBimestres = spreadsheet.getSheetByName("Bimestres");

  if (!paginaBimestres) {
    throw new Error("Página 'Bimestres' não encontrada na planilha.");
  }

  const valores = paginaBimestres.getDataRange().getValues();
  LimparDados();
  for (let i = 1; i < valores.length; i++) { // Comece da segunda linha
    const bimestre = valores[i][0]; // Coluna Planilhas
    const link = valores[i][1]; // Coluna Link

    if (link) { // Verifica se ambas as células não estão vazias
      Lista_Turmas(bimestre); // Chame Lista_Turmas passando a planilha
    }
  }
}

function LimparDados() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const paginasParaLimpar = ["Consolidado Notas", "Consolidado Trilhas", "Consolidado Atitudinal"];

  for (const nomePagina of paginasParaLimpar) {
    const pagina = spreadsheet.getSheetByName(nomePagina);

    if (pagina) {
      const ultimaLinha = pagina.getLastRow();
      if (ultimaLinha > 1) { // Verifica se há mais de uma linha
        pagina.getRange(2, 1, ultimaLinha - 1, pagina.getLastColumn()).clearContent();
      } // Se não houver mais linhas, não faz nada (evita erro)
    } // Se a página não existir, não faz nada (evita erro)
  }
}

function Lista_Turmas(pbimestre) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Acessar a página "Bimestre"
  const paginaBimestre = spreadsheet.getSheetByName('Bimestres');

  // Encontrar a linha com o bimestre correspondente
  const valoresBimestre = paginaBimestre.getRange(2, 1, paginaBimestre.getLastRow() - 1, 1).getValues(); // Coluna A (Bimestre) a partir da segunda linha
  const linhaBimestre = valoresBimestre.findIndex(valor => valor[0] == pbimestre) + 2; // +2 para ajustar o índice da linha

  if (linhaBimestre > 1) { // Verifica se a linha foi encontrada
    const linkPlanilha = paginaBimestre.getRange(linhaBimestre, 2).getValue(); // Coluna B (Link)

    if (linkPlanilha) {
      PlanilhaBimestre = SpreadsheetApp.openByUrl(linkPlanilha);
    } else {
      SpreadsheetApp.getUi().alert('Link não encontrado para o bimestre ' + pbimestre);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert('Bimestre ' + pbimestre + ' não encontrado');
    return;
  }

  const todasAsPaginas = PlanilhaBimestre.getSheets();

  for (const pagina of todasAsPaginas) {
    const nomePagina = pagina.getName();

    // Verificar se o nome da página começa com um número
    if (/^[6-9]/.test(nomePagina)) { 
      consolidarNotasEF(PlanilhaBimestre,nomePagina);
      consolidarAtitudinal(PlanilhaBimestre,nomePagina,8);
    }
    if (/^[1-3]/.test(nomePagina)) { 
      consolidarNotasEM(PlanilhaBimestre,nomePagina);
      consolidarAtitudinal(PlanilhaBimestre,nomePagina,20);
    }
  }
}

function consolidarNotasEF(PlanilhaBimestre, nomePagina) {
  // 1. Nomes das colunas na página de notas do EF.
  const nomesDasNotas = RetornaColunasEF(PlanilhaBimestre,nomePagina);

  // 2. Acessar a planilha e as páginas
  const spreadsheetDestino = SpreadsheetApp.getActiveSpreadsheet();
  const paginaNotas = PlanilhaBimestre.getSheetByName(nomePagina);
  const paginaConsolidado = spreadsheetDestino.getSheetByName('Consolidado Notas');

  // 3. Obter os dados da página "Notas"
  const dadosNotas = paginaNotas.getDataRange().getValues();

  // 4. Obter informações da turma e bimestre
  const anoTurma = dadosNotas[1][0]; // Célula A2
  const turma = anoTurma.slice(-1); // Última letra
  const ano = anoTurma.slice(0, -1);
  const bimestre = dadosNotas[0][1]; // Célula B1

  // 5. Processar os dados e consolidar as notas
  const dadosConsolidados = [];
  for (let i = 3; i < dadosNotas.length; i++) {
    const aluno = dadosNotas[i][0];
    if (!aluno) break; // Parar se o nome do aluno for vazio

    for (let j = 1; j < 9; j++) {
      const materia = nomesDasNotas[j - 1];
      const nota = dadosNotas[i][j];

      dadosConsolidados.push([ano, turma, bimestre, aluno, materia, nota]);
    }
  }


 // 6. Encontrar a primeira linha livre e inserir os dados
  const ultimaLinha = paginaConsolidado.getLastRow();
  const primeiraLinhaVazia = ultimaLinha + 1;

  paginaConsolidado.getRange(primeiraLinhaVazia, 1, dadosConsolidados.length, 6).setValues(dadosConsolidados);

}

function RetornaColunasEM(PlanilhaBimestre,NomePagina) {
  
  const sheet = PlanilhaBimestre.getSheetByName(NomePagina);

  if (!sheet) {
    throw new Error(`Página "${NomePagina}" não encontrada.`);
  }

  const range = sheet.getRange("B2:U2");
  const values = range.getValues();

  // Converte a matriz 2D em um vetor 1D
  return values[0]; 
}

function RetornaColunasEF(PlanilhaBimestre,NomePagina) {
  
  const sheet = PlanilhaBimestre.getSheetByName(NomePagina);

  if (!sheet) {
    throw new Error(`Página "${NomePagina}" não encontrada.`);
  }

  const range = sheet.getRange("B2:I2");
  const values = range.getValues();

  // Converte a matriz 2D em um vetor 1D
  return values[0]; 
}

function consolidarNotasEM(PlanilhaBimestre,nomePagina) {
  // 1. Nomes das colunas na página de notas EM.
  const nomesDasNotas = RetornaColunasEM(PlanilhaBimestre,nomePagina);

  // 2. Acessar a planilha e as páginas
  const paginaNotas = PlanilhaBimestre.getSheetByName(nomePagina);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const paginaConsolidadoNotas = spreadsheet.getSheetByName('Consolidado Notas');
  const paginaConsolidadoTrilhas = spreadsheet.getSheetByName('Consolidado Trilhas');

  // 3. Obter os dados da página "Notas"
  const dadosNotas = paginaNotas.getDataRange().getValues();

  // 4. Obter informações da turma e bimestre
  const anoTurma = dadosNotas[1][0]; // Célula A2
  const turma = anoTurma.slice(-1); // Última letra
  const ano = anoTurma.slice(0, -1);
  const bimestre = dadosNotas[0][1]; // Célula B1

  // 5. Processar os dados e consolidar as notas
  const dadosConsolidadosNotas = [];
  const dadosConsolidadosTrilhas = [];
  for (let i = 3; i < dadosNotas.length; i++) {
    const aluno = dadosNotas[i][0];
    if (!aluno) break; // Parar se o nome do aluno for vazio

    for (let j = 1; j < 21; j++) {
      const materia = nomesDasNotas[j - 1];
      const nota = dadosNotas[i][j];
      if(!isNaN(nota)) {
        if (nota !== "") dadosConsolidadosNotas.push([ano, turma, bimestre, aluno, materia, nota]);
      }
      else {
        dadosConsolidadosTrilhas.push([ano, turma, bimestre, aluno, materia, nota]);
      }
    }
  }


 // 6. Encontrar a primeira linha livre e inserir os dados
  const ultimaLinhaNotas = paginaConsolidadoNotas.getLastRow();
  const primeiraLinhaVaziaNotas = ultimaLinhaNotas + 1;
  paginaConsolidadoNotas.getRange(primeiraLinhaVaziaNotas, 1, dadosConsolidadosNotas.length, 6).setValues(dadosConsolidadosNotas);

  const ultimaLinhaTrilhas = paginaConsolidadoTrilhas.getLastRow();
  const primeiraLinhaVaziaTrilhas = ultimaLinhaTrilhas + 1;
  paginaConsolidadoTrilhas.getRange(primeiraLinhaVaziaTrilhas, 1, dadosConsolidadosTrilhas.length, 6).setValues(dadosConsolidadosTrilhas);


}


function consolidarAtitudinal(PlanilhaBimestre,nomePagina,posicaoatitudes) {
  // 1. Configuração (substitua pelos valores reais)
  const nomesDasAtitudes = ['FALTOSO','NÃO FAZ AS ATIVIDADES',	'DESATENTO',	'DIFICULDADE',	'SATISFATÓRIO','NECESSIDADE DE ATENÇÃO ESPECIAL/AEE','OBSERVAÇÃO']; 

  // 2. Acessar a planilha e as páginas
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const paginaNotas = PlanilhaBimestre.getSheetByName(nomePagina);
  const paginaConsolidado = spreadsheet.getSheetByName('Consolidado Atitudinal');

  // 3. Obter os dados da página nomePagina
  const dadosNotas = paginaNotas.getDataRange().getValues();

  // 4. Obter informações da turma e bimestre
  const anoTurma = dadosNotas[1][0]; // Célula A2
  const turma = anoTurma.slice(-1); // Última letra
  const ano = anoTurma.slice(0, -1);
  const bimestre = dadosNotas[0][1]; // Célula B1

  // 5. Processar os dados e consolidar as notas
  const dadosConsolidados = [];
  for (let i = 3; i < dadosNotas.length; i++) {
    const aluno = dadosNotas[i][0];
    if (!aluno) break; // Parar se o nome do aluno for vazio

    for (let j = 1; j < 8; j++) {
      const aspecto = nomesDasAtitudes[j - 1];
      if(dadosNotas[i][j+posicaoatitudes]==false)
        var parecer = "NÃO";
      else if(dadosNotas[i][j+posicaoatitudes]==true)
        var parecer = "SIM";
      else
        var parecer = dadosNotas[i][j+posicaoatitudes];

      dadosConsolidados.push([ano, turma, bimestre, aluno, aspecto, parecer]);
    }
  }


 // 6. Encontrar a primeira linha livre e inserir os dados
  const ultimaLinha = paginaConsolidado.getLastRow();
  const primeiraLinhaVazia = ultimaLinha + 1;

  paginaConsolidado.getRange(primeiraLinhaVazia, 1, dadosConsolidados.length, 6).setValues(dadosConsolidados);

}

