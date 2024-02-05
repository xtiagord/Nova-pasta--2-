let atendentes = [
    { nome: 'Atendente 1', atendimentos: [] },
    { nome: 'Atendente 2', atendimentos: [] },
    { nome: 'Atendente 3', atendimentos: [] },
    // Adicione outros atendentes conforme necessário
  ];

  let indiceAtual = 0;

  function proximoAtendente() {
    const atendenteAnterior = atendentes[indiceAtual];
    if (atendenteAnterior) {
      const dataAtual = new Date();
      const dataFormatada = `${dataAtual.toLocaleDateString()} às ${dataAtual.toLocaleTimeString()}`;
      atendenteAnterior.atendimentos.push(dataFormatada);
    }

    if (indiceAtual < atendentes.length - 1) {
      indiceAtual++;
    } else {
      indiceAtual = 0;
    }

    mostrarAtendenteAtual();
  }

  function mostrarAtendenteAtual() {
    document.getElementById('atendenteAtual').innerText = atendentes[indiceAtual].nome;
    // Limpar a exibição do horário quando o atendente muda
    document.getElementById('infoAtendimento').innerText = 'Nenhum';
  }

  function exibirAtendimento(indice) {
    const atendente = atendentes[indice - 1];
    const atendimentosAnteriores = atendente.atendimentos;
    const listaAtendimentos = document.getElementById('listaAtendimentos');

    // Limpar a lista antes de preenchê-la novamente
    listaAtendimentos.innerHTML = '';

    if (atendimentosAnteriores.length > 0) {
      for (const atendimento of atendimentosAnteriores) {
        const listItem = document.createElement('li');
        listItem.classList.add('list-group-item');
        listItem.innerText = atendimento;
        listaAtendimentos.appendChild(listItem);
      }
    } else {
      const listItem = document.createElement('li');
      listItem.classList.add('list-group-item');
      listItem.innerText = 'Nenhum atendimento anterior.';
      listaAtendimentos.appendChild(listItem);
    }

    // Exibir o modal
    new bootstrap.Modal(document.getElementById('modalAtendimentos')).show();
  }

  function exportarParaExcel() {
  const data = atendentes.map(atendente => ({
    'Atendente': atendente.nome,
    'Atendimentos Anteriores': atendente.atendimentos.join(', ')
  }));

  XlsxPopulate.fromBlankAsync().then(wb => {
    const ws = wb.sheet(0);

    // Adicionando estilo ao cabeçalho (A1)
    ws.cell('A1').style({
      fill: { type: 'solid', color: '00008B' }, // Azul 
      fontColor: 'FFFFFF', // Branco
      bold: true
    });

    // Adicionando estilo à célula A2
    ws.cell('B1').style({
      fill: { type: 'solid', color: 'FFFF00' }, // Amarelo
      bold: true,
      fontColor: '000000', // Black
    });

    // Preenchendo os dados
    ws.cell('A1').value('Atendente');
    ws.cell('B1').value('Atendimentos Anteriores');
    
    for (let i = 0; i < data.length; i++) {
      ws.cell(`A${i + 2}`).value(data[i]['Atendente']);
      ws.cell(`B${i + 2}`).value(data[i]['Atendimentos Anteriores']);
    }

    const larguraColunaA = Math.max(...data.map(item => item['Atendente'].length)) + 2; // +2 para dar espaço extra
    const larguraColunaB = Math.max(...data.map(item => item['Atendimentos Anteriores'].length)) + 2; // +2 para dar espaço extra

    ws.column('A').width(larguraColunaA);
    ws.column('B').width(larguraColunaB);

    // Exportar o arquivo Excel
    wb.outputAsync().then(blob => {
      saveAs(blob, 'atendentes.xlsx');
    });
  });
}
