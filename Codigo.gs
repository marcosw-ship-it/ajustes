/**
 * SISTEMA IBRACEL - VFINAL (LEITOR INTELIGENTE PARA MÚLTIPLOS CÓDIGOS COM VÍRGULA)
 */

const COLUNAS_FIXAS = ['ID', 'ALUNO', 'TURMA', 'CPF', 'OBSERVAÇÕES', 'NOTAS', 'TOTAL', 'SITUAÇÃO', 'MEDIA', 'COMPROVANTE DE INSCRIÇÃO']; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('Frequência IBRACEL');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function formatarNome(nome) {
  return nome.replace(/\w\S*/g, function(txt) { return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase(); });
}

function getTurmas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Grupos'); 
    if (!sheet) sheet = ss.getSheetByName('Turma');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const bg = sheet.getDataRange().getBackgrounds();
    const turmas = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        turmas.push({ nome: data[i][0].toString().trim(), cor: bg[i][1] || '#229783' });
      }
    }
    return turmas; 
  } catch (e) { return []; }
}

function getColunasFrequencia() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BANCO DE DADOS');
    if (!sheet) return [];
    const lastCol = sheet.getLastColumn();
    if (lastCol < 5) return []; 
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    return headers.map((h, i) => ({ nome: h ? h.toString().trim() : '', index: i }))
                  .filter(h => h.index >= 4 && h.nome !== '' && !COLUNAS_FIXAS.includes(h.nome.toUpperCase()));
  } catch (e) { return []; }
}

function getTodasPessoas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BANCO DE DADOS');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const colunas = getColunasFrequencia();
    const pessoas = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() !== '') {
        let registro = {
          id: data[i][0] ? data[i][0].toString() : '',
          aluno: data[i][1].toString(),
          turma: data[i][2] ? data[i][2].toString() : '',
          cpf: data[i][3] ? data[i][3].toString() : '', 
          frequencias: {}
        };
        colunas.forEach(col => {
          registro.frequencias[col.nome] = data[i][col.index] ? data[i][col.index].toString() : 'Faltou';
        });
        pessoas.push(registro);
      }
    }
    return pessoas;
  } catch (e) { return []; }
}

function adicionarPessoa(aluno, turma, cpf) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BANCO DE DADOS');
    
    if (!sheet) return { success: false, message: "Planilha BANCO DE DADOS não encontrada." };

    const nomeFormatado = formatarNome(aluno.trim());
    const cpfInput = cpf ? cpf.trim() : "";
    const data = sheet.getDataRange().getValues();
    
    // 1. Verificação de Duplicados/Conflitos
    let matches = [];
    let jaExisteNestaTurma = false;
    for (let i = 1; i < data.length; i++) {
      const nomePlanilha = data[i][1] ? data[i][1].toString().toLowerCase().trim() : "";
      const cpfPlanilha = data[i][3] ? data[i][3].toString().trim() : "";
      
      if (nomePlanilha === nomeFormatado.toLowerCase()) {
        if (cpfInput !== "" && cpfPlanilha !== "" && cpfInput !== cpfPlanilha) continue;
        if (data[i][2].toString().trim() === turma) { jaExisteNestaTurma = true; break; }
        matches.push({ id: data[i][0], turma: data[i][2], cpf: cpfPlanilha || "Sem CPF cadastrado" });
      }
    }

    if (jaExisteNestaTurma) return { success: false, message: "Este aluno já está cadastrado nesta turma!" };

    if (matches.length > 0) {
      return { 
        success: false, conflict: true, 
        message: `Encontramos registros com o nome "${nomeFormatado}". É a mesma pessoa?`, 
        matches: matches, aluno: nomeFormatado, destino: turma 
      };
    }

    // 2. Lógica de Inserção Corrigida
    const id = Utilities.getUuid();
    const totalColunas = sheet.getLastColumn();
    const colunasFreq = getColunasFrequencia(); // Obtém apenas colunas de turno
    
    // Cria um array vazio com o tamanho exato da planilha
    let novaLinha = new Array(totalColunas).fill("");
    
    // Preenche os dados fixos
    novaLinha[0] = id;           // Coluna ID
    novaLinha[1] = nomeFormatado; // Coluna ALUNO
    novaLinha[2] = turma;         // Coluna TURMA
    novaLinha[3] = cpfInput;      // Coluna CPF

    // Preenche "Faltou" APENAS nas colunas que o sistema reconhece como frequência
    colunasFreq.forEach(col => {
      if (col.index < totalColunas) {
        novaLinha[col.index] = 'Faltou';
      }
    });

    sheet.appendRow(novaLinha);
    SpreadsheetApp.flush();
    
    return { success: true, message: "Adicionado com sucesso!" };

  } catch(e) { 
    return { success: false, message: e.toString() }; 
  } finally { 
    lock.releaseLock();
  }
}

function moverPessoa(aluno, origem, destino, acao, idExistente, cpfInput) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BANCO DE DADOS');
    const nome = formatarNome(aluno);
    const data = sheet.getDataRange().getValues();

    if (acao === 'mover') {
      for(let i=1; i<data.length; i++) {
        if(data[i][0] == idExistente && data[i][2] == origem) {
          sheet.getRange(i+1, 3).setValue(destino); 
          SpreadsheetApp.flush();
          return { success: true, message: "Aluno movido para " + destino };
        }
      }
    } else if (acao === 'manter') {
      let cpfAntigo = "";
      for(let i=1; i<data.length; i++) { if(data[i][0] == idExistente) { cpfAntigo = data[i][3]; break; } }
      const row = [idExistente, nome, destino, cpfAntigo]; 
      const total = sheet.getLastColumn();
      for(let i=4; i<total; i++) row.push('Faltou');
      sheet.appendRow(row);
      SpreadsheetApp.flush();
      return { success: true, message: "Aluno espelhado em ambas as turmas!" };
    } else if (acao === 'novo') {
      const row = [Utilities.getUuid(), nome, destino, cpfInput || ""]; 
      const total = sheet.getLastColumn();
      for(let i=4; i<total; i++) row.push('Faltou');
      sheet.appendRow(row);
      SpreadsheetApp.flush();
      return { success: true, message: "Novo aluno homônimo criado com sucesso!" };
    }
    return { success: false, message: "Erro ao processar conflito." };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function atualizarFrequenciaIndividual(identificador, turnoNome, valor) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BANCO DE DADOS');
    const data = sheet.getDataRange().getValues();
    
    let colIndex = -1;
    for (let i = 0; i < data[0].length; i++) {
      if (data[0][i].toString().trim().toLowerCase() === turnoNome.trim().toLowerCase()) {
        colIndex = i + 1; break;
      }
    }
    
    if (colIndex <= 0) return { success: false, message: "Turno não encontrado na planilha." };

    const codigosBusca = identificador.toString().toLowerCase().split(/[\n,;|\/]+/).map(c => c.trim()).filter(c => c !== '');
    if (codigosBusca.length === 0) return { success: false, message: "Nenhum código recebido." };

    let count = 0;
    
    for (let i = 1; i < data.length; i++) {
      let dbIds = data[i][0].toString().toLowerCase().split(/[\n,;|\/]+/).map(c => c.trim());
      let dbCpfs = (data[i][3] || "").toString().toLowerCase().split(/[\n,;|\/]+/).map(c => c.trim());
      
      let codigosDaPessoa = dbIds.concat(dbCpfs);
      let match = codigosBusca.some(cod => codigosDaPessoa.includes(cod));

      if (match) {
        sheet.getRange(i + 1, colIndex).setValue(valor);
        count++;
      }
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: `Salvo!` };
  } catch (e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function editarNomeAlunoPorId(id, nomeNovo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BANCO DE DADOS');
    const data = sheet.getDataRange().getValues();
    const novo = formatarNome(nomeNovo.trim());
    
    for(let i=1; i<data.length; i++) {
      if(data[i][0].toString() === id.toString()) {
        sheet.getRange(i+1, 2).setValue(novo);
      }
    }
    SpreadsheetApp.flush();
    return { success: true, message: "Nome alterado com sucesso!" };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function removerPessoa(id, aluno) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BANCO DE DADOS');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0].toString() === id.toString()) { 
      sheet.deleteRow(i+1); 
      SpreadsheetApp.flush(); 
      return { success: true, message: "Removido!" }; 
    }
  }
  return { success: false, message: "Não encontrado." };
}

function carregarMetasTurmas() {
  try { return JSON.parse(PropertiesService.getUserProperties().getProperty('metasTurmas') || "{}"); } catch(e) { return {}; }
}
function salvarMetasTurmas(m) { PropertiesService.getUserProperties().setProperty('metasTurmas', JSON.stringify(m)); return {success:true}; }
function getTurmasComEsperados() { 
  const t = getTurmas(); 
  const m = carregarMetasTurmas();
  return t.map(turma => ({...turma, esperados: m[turma.nome] || 0}));
}

function gerarRelatorioAnalytics() {
  try {
    const pessoas = getTodasPessoas();
    const turmas = getTurmas();
    const metas = carregarMetasTurmas();
    const colunas = getColunasFrequencia();
    
    let gruposAtivos = new Set();
    colunas.forEach(c => {
       let n = c.nome.toLowerCase();
       if(n.includes('manh')) gruposAtivos.add('Manhã');
       else if(n.includes('tard')) gruposAtivos.add('Tarde');
       else if(n.includes('noit')) gruposAtivos.add('Noite');
       else gruposAtivos.add(c.nome); 
    });
    
    let listaGrupos = Array.from(gruposAtivos);
    const ordemPrioridade = ['manhã', 'tarde', 'noite'];
    listaGrupos.sort((a, b) => {
      let iA = ordemPrioridade.indexOf(a.toLowerCase());
      let iB = ordemPrioridade.indexOf(b.toLowerCase());
      if (iA !== -1 && iB !== -1) return iA - iB;
      if (iA !== -1) return -1;
      if (iB !== -1) return 1;
      return 0; 
    });

    const stats = turmas.map(t => {
      const alunos = pessoas.filter(p => p.turma === t.nome);
      let contagemPorGrupo = {};
      listaGrupos.forEach(g => contagemPorGrupo[g] = 0);

      const listaAlunos = alunos.map(a => {
        let freqFormatada = {};
        Object.keys(a.frequencias).forEach(col => { freqFormatada[col] = a.frequencias[col] === 'Presente' ? 'P' : 'F'; });
        return { nome: a.aluno, frequencias: freqFormatada };
      });

      alunos.forEach(a => { 
        let presencasUnicasDoAluno = {};
        Object.keys(a.frequencias).forEach(k => {
          if(a.frequencias[k] === 'Presente') {
            let lowK = k.toLowerCase();
            let g = k;
            if(lowK.includes('manh')) g = 'Manhã';
            else if(lowK.includes('tard')) g = 'Tarde';
            else if(lowK.includes('noit')) g = 'Noite';
            presencasUnicasDoAluno[g] = true;
          }
        });
        Object.keys(presencasUnicasDoAluno).forEach(g => { if(presencasUnicasDoAluno[g]) contagemPorGrupo[g]++; });
      });
      
      const esperados = Number(metas[t.nome] || alunos.length);
      let graficosTurma = [];
      let maxPresentes = 0;

      listaGrupos.forEach(g => {
        let p = contagemPorGrupo[g];
        if (p > maxPresentes) maxPresentes = p;
        let perc = esperados > 0 ? ((p / esperados) * 100).toFixed(1) : 0;
        graficosTurma.push({ nomeGrupo: g, presentes: p, perc: Number(perc) });
      });

      return {
        turma: t.nome, cor: t.cor, esperados: esperados, totalAlunos: alunos.length,
        percGeral: (esperados > 0 ? ((maxPresentes/esperados)*100).toFixed(1) : 0),
        graficos: graficosTurma.filter(g => g.presentes >= 0),
        listaAlunos: listaAlunos
      };
    });

    let destaquesPorGrupo = {};
    listaGrupos.forEach(g => {
        let turmasComOGrupo = stats.filter(t => t.esperados > 0 && t.graficos.find(gr => gr.nomeGrupo === g && gr.presentes > 0));
        if(turmasComOGrupo.length > 0) {
            turmasComOGrupo.sort((a,b) => {
                let pA = a.graficos.find(gr => gr.nomeGrupo === g).perc;
                let pB = b.graficos.find(gr => gr.nomeGrupo === g).perc;
                return pB - pA;
            });
            let best = turmasComOGrupo[0];
            let gData = best.graficos.find(gr => gr.nomeGrupo === g);
            destaquesPorGrupo[g] = { turma: best.turma, perc: gData.perc, presentes: gData.presentes, esperados: best.esperados };
        }
    });

    const turmasValidas = stats.filter(t => t.esperados > 0);
    const atencao = [];
    stats.forEach(t => {
       if (t.esperados === 0) return;
       let turnosBaixos = [];
       t.graficos.forEach(g => { if(g.perc < 70) turnosBaixos.push(`${g.nomeGrupo}: ${g.perc}%`); });
       if(turnosBaixos.length > 0) atencao.push({ turma: t.turma, details: turnosBaixos.join(' | ') });
    });

    return { success: true, relatorio: { statsPorTurma: stats, colunas: colunas, listaGrupos: listaGrupos, destaquesPorGrupo: destaquesPorGrupo, turmasAtencao: atencao } };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getHeaderHTML(dados, titulo) {
  let dataFmt = dados.data ? dados.data.split('-').reverse().join('/') : '___/___/___';
  let logo1 = dados.logo1 ? `<img src="${dados.logo1}" style="height:60px; max-width:80px; object-fit:contain; margin-left:10px;">` : '';
  let logo2 = dados.logo2 ? `<img src="${dados.logo2}" style="height:60px; max-width:80px; object-fit:contain; margin-left:10px;">` : '';
  return `
    <table style="width:100%; border-bottom:3px solid #229783; margin-bottom:20px; border-collapse:collapse; border:none;">
       <tr>
         <td style="width:25%; border:none; text-align:left; vertical-align:middle; padding-bottom:10px;"><img src="https://i.postimg.cc/W3QnQkr3/Prancheta-2.png" style="height:70px;"></td>
         <td style="width:50%; text-align:center; border:none; vertical-align:middle; padding-bottom:10px;">
            <h2 style="margin:0; color:#1a5151; font-family:Arial; font-size:18px;">${titulo}</h2>
            <p style="margin:5px 0 0 0; color:#666; font-family:Arial; font-size:12px;">Data: ${dataFmt} | Município: ${dados.municipio || '___'}</p>
         </td>
         <td style="width:25%; text-align:right; border:none; vertical-align:middle; white-space:nowrap; padding-bottom:10px;">${logo1}${logo2}</td>
       </tr>
    </table>`;
}

function gerarPDF(dados) {
  try {
    const pessoas = getTodasPessoas();
    const turmas = getTurmas();
    let html = `<html><head><style>body{font-family:Arial; -webkit-print-color-adjust:exact; print-color-adjust:exact;} table{width:100%;border-collapse:collapse;} td,th{border:1px solid #ccc;padding:8px;} .break{page-break-before:always;}</style></head><body>`;
    turmas.forEach((t, i) => {
      if(i>0) html += '<div class="break"></div>';
      html += getHeaderHTML(dados, "LISTA DE ASSINATURA");
      
      let corFonte = getCorTexto(t.cor); // AQUI ESTAVA O ERRO DE SINTAXE QUE CONSERTEI!
      
      html += `<h2 style="background-color:${t.cor} !important; color:${corFonte} !important; padding:10px; border-radius:5px; font-weight:bold; font-size:16px; margin-bottom:10px; -webkit-print-color-adjust: exact;">Turma: ${t.nome}</h2>
               <table><thead><tr style="background-color:#1a5151 !important; color:white !important; -webkit-print-color-adjust: exact;"><th width="5%">#</th><th width="45%">Nome do Aluno</th><th width="50%">Assinatura</th></tr></thead><tbody>`;
      const alunos = pessoas.filter(p => p.turma === t.nome).sort((a,b)=>a.aluno.localeCompare(b.aluno));
      alunos.forEach((a, idx) => { html += `<tr><td style="text-align:center;">${idx+1}</td><td>${a.aluno}</td><td></td></tr>`; });
      html += `</tbody></table>`;
    });
    html += `</body></html>`;
    return { success: true, html: html };
  } catch(e) { return { success: false, message: e.toString() }; }
}

function gerarPDFCompacto(dados) {
  try {
    const pessoas = getTodasPessoas();
    const turmas = getTurmas();
    const colunas = getColunasFrequencia();
    let html = `<html><head><style>
      body{font-family:Arial; font-size:12px; -webkit-print-color-adjust:exact; print-color-adjust:exact;} 
      table{width:100%;border-collapse:collapse; margin-bottom:20px;} 
      td,th{border:1px solid #ddd;padding:6px;} 
      .break{page-break-before:always;}
      .sq-p { display:inline-block; width:18px; height:18px; background-color:#27ae60 !important; color:white !important; text-align:center; font-size:11px; line-height:18px; font-weight:bold; border-radius:3px; }
      .sq-f { display:inline-block; width:18px; height:18px; background-color:#e74c3c !important; color:white !important; text-align:center; font-size:11px; line-height:18px; font-weight:bold; border-radius:3px; }
    </style></head><body>`;
    
    turmas.forEach((t, i) => {
      if(i>0) html += '<div class="break"></div>';
      html += getHeaderHTML(dados, "RELATÓRIO GERAL DE FREQUÊNCIA");
      
      let corFonte = getCorTexto(t.cor);
      
      html += `<div style="background-color:${t.cor} !important; color:${corFonte} !important; padding:8px; border-radius:3px; font-weight:bold; margin-bottom:5px;">Turma: ${t.nome}</div>
               <table><thead><tr style="background-color:#1a5151 !important; color:white !important; -webkit-print-color-adjust: exact;">
               <th style="width:50%; text-align:left;">Nome do Aluno</th>`;
      colunas.forEach(c => html += `<th style="text-align:center;">${c.nome}</th>`);
      html += `</tr></thead><tbody>`;
      
      const alunos = pessoas.filter(p => p.turma === t.nome).sort((a,b)=>a.aluno.localeCompare(b.aluno));
      alunos.forEach((a) => { 
        html += `<tr><td style="font-weight:bold; color:#444;">${a.aluno}</td>`;
        colunas.forEach(c => {
           const isP = a.frequencias[c.nome] === 'Presente';
           html += `<td style="text-align:center;">${isP ? '<div class="sq-p">P</div>' : '<div class="sq-f">F</div>'}</td>`;
        });
        html += `</tr>`;
      });
      html += `</tbody></table>`;
    });
    html += `</body></html>`;
    return { success: true, html: html };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function salvarLogos(logos) { PropertiesService.getUserProperties().setProperty('logos', JSON.stringify(logos)); }
function carregarLogos() { const l = PropertiesService.getUserProperties().getProperty('logos'); return l ? JSON.parse(l) : { logo1: null, logo2: null }; }
function salvarConfiguracoes(c) { PropertiesService.getUserProperties().setProperty('config', JSON.stringify(c)); }
function carregarConfiguracoes() { const c = PropertiesService.getUserProperties().getProperty('config'); return c ? JSON.parse(c) : { data: '', municipio: '' }; }

// ==============================================================
// LEITOR AUTOMÁTICO DO APPSHEET PARA O BANCO DE DADOS
// ==============================================================
function processarLeiturasAppSheet() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLeituras = ss.getSheetByName('Leituras');
    const sheetBanco = ss.getSheetByName('BANCO DE DADOS');
    
    if (!sheetLeituras || !sheetBanco) return;
    
    const dadosLeituras = sheetLeituras.getDataRange().getValues();
    const dadosBanco = sheetBanco.getDataRange().getValues();
    
    const cabecalhosBanco = dadosBanco[0].map(c => c.toString().trim().toLowerCase());
    
    let alterouBanco = false;
    
    for (let i = 1; i < dadosLeituras.length; i++) {
      let turnoLido = dadosLeituras[i][1] ? dadosLeituras[i][1].toString().trim() : "";
      let codigoLido = dadosLeituras[i][2] ? dadosLeituras[i][2].toString().trim() : "";
      let status = dadosLeituras[i][3] ? dadosLeituras[i][3].toString().trim() : "";
      
      if (turnoLido !== "" && codigoLido !== "" && status === "") {
        
        let colTurnoIndex = cabecalhosBanco.indexOf(turnoLido.toLowerCase());
        
        if (colTurnoIndex > -1) {
          let codigosBusca = codigoLido.toLowerCase().split(/[\n,;|\/]+/).map(c => c.trim()).filter(c => c !== '');
          
          for (let r = 1; r < dadosBanco.length; r++) {
            let dbIds = dadosBanco[r][0].toString().toLowerCase().split(/[\n,;|\/]+/).map(c => c.trim());
            let dbCpfs = (dadosBanco[r][3] || "").toString().toLowerCase().split(/[\n,;|\/]+/).map(c => c.trim());
            
            let codigosDaPessoa = dbIds.concat(dbCpfs);
            let match = codigosBusca.some(cod => codigosDaPessoa.includes(cod));
            
            if (match) {
              sheetBanco.getRange(r + 1, colTurnoIndex + 1).setValue("Presente");
              alterouBanco = true;
            }
          }
        }
        sheetLeituras.getRange(i + 1, 4).setValue("Processado ✅");
      }
    }
    
    if (alterouBanco) {
      SpreadsheetApp.flush();
    }
    
  } catch (e) {
  } finally {
    lock.releaseLock();
  }
}

// ==============================================================
// SISTEMA DE CONTRASTE AUTOMÁTICO (YIQ)
// ==============================================================
function getCorTexto(hexColor) {
  if (!hexColor) return '#ffffff'; 
  
  let hex = hexColor.replace("#", "");
  
  if (hex.length === 3) {
    hex = hex.split('').map(h => h + h).join('');
  }
  
  let r = parseInt(hex.substr(0, 2), 16);
  let g = parseInt(hex.substr(2, 2), 16);
  let b = parseInt(hex.substr(4, 2), 16);
  
  let yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000;
  
  return (yiq >= 128) ? '#1a1a1a' : '#ffffff'; 
}
