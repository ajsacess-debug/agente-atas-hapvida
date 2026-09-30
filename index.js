const DocxBuilder = require('docx').Document;
const { Packer, Paragraph, Table, TableRow, TableCell, TextRun, BorderStyle, WidthType, AlignmentType } = require('docx');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method === 'GET') {
    res.status(200).json({ status: 'Agente de Atas funcionando!' });
    return;
  }

  if (req.method === 'POST') {
    try {
      const dados = req.body;

      const doc = new DocxBuilder({
        sections: [{
          children: [
            new Paragraph({
              text: 'PÚBLICO',
              alignment: AlignmentType.CENTER,
              run: new TextRun({ size: 28, bold: true, color: 'E31D23' })
            }),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: 'ATA DE REUNIÃO',
              alignment: AlignmentType.CENTER,
              run: new TextRun({ size: 36, bold: true, color: '1F4E78' })
            }),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: (dados.assunto || 'ASSUNTO').toUpperCase(),
              alignment: AlignmentType.CENTER,
              run: new TextRun({ size: 22, bold: true })
            }),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: `REALIZADA EM ${(dados.data || 'DATA').toUpperCase()}.`,
              alignment: AlignmentType.CENTER,
              run: new TextRun({ size: 22, bold: true })
            }),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: 'Objetivo da Reunião',
              run: new TextRun({ size: 24, bold: true })
            }),
            new Paragraph({ text: dados.objetivo || '' }),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: 'Resumo Executivo',
              run: new TextRun({ size: 24, bold: true })
            }),
            new Paragraph({ text: dados.resumo || '' }),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: 'Principais Definições e Entendimentos',
              run: new TextRun({ size: 24, bold: true })
            }),
            ...(dados.definicoes || []).flatMap(def => [
              new Paragraph({
                text: def.topico || '',
                run: new TextRun({ size: 22, bold: true })
              }),
              ...((def.pontos || []).map(p => new Paragraph({ text: p, bullet: { level: 0 } })))
            ]),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: 'Riscos Identificados',
              run: new TextRun({ size: 24, bold: true })
            }),
            ...(dados.riscos?.altoImpacto ? [
              new Paragraph({
                text: 'Alto Impacto',
                run: new TextRun({ size: 22, bold: true })
              }),
              ...(dados.riscos.altoImpacto.map(r => new Paragraph({ text: r, bullet: { level: 0 } })))
            ] : []),
            ...(dados.riscos?.medioImpacto ? [
              new Paragraph({
                text: 'Médio Impacto',
                run: new TextRun({ size: 22, bold: true })
              }),
              ...(dados.riscos.medioImpacto.map(r => new Paragraph({ text: r, bullet: { level: 0 } })))
            ] : []),
            new Paragraph({ text: '' }),
            new Paragraph({
              text: 'Próximos Passos',
              run: new TextRun({ size: 24, bold: true })
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  cells: [
                    new TableCell({
                      width: { size: 70, type: WidthType.PERCENTAGE },
                      borders: { top: { style: BorderStyle.SINGLE, size: 6 }, bottom: { style: BorderStyle.SINGLE, size: 6 }, left: { style: BorderStyle.SINGLE, size: 6 }, right: { style: BorderStyle.SINGLE, size: 6 } },
                      shading: { type: 'clear', fill: 'D9D9D9' },
                      children: [new Paragraph({ text: 'Ação', run: new TextRun({ bold: true }) })]
                    }),
                    new TableCell({
                      width: { size: 30, type: WidthType.PERCENTAGE },
                      borders: { top: { style: BorderStyle.SINGLE, size: 6 }, bottom: { style: BorderStyle.SINGLE, size: 6 }, left: { style: BorderStyle.SINGLE, size: 6 }, right: { style: BorderStyle.SINGLE, size: 6 } },
                      shading: { type: 'clear', fill: 'D9D9D9' },
                      children: [new Paragraph({ text: 'Responsável', run: new TextRun({ bold: true }) })]
                    })
                  ]
                }),
                ...((dados.proximosPassos || []).map(p =>
                  new TableRow({
                    cells: [
                      new TableCell({ width: { size: 70, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.SINGLE, size: 6 }, bottom: { style: BorderStyle.SINGLE, size: 6 }, left: { style: BorderStyle.SINGLE, size: 6 }, right: { style: BorderStyle.SINGLE, size: 6 } }, children: [new Paragraph({ text: p.acao || '' })] }),
                      new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.SINGLE, size: 6 }, bottom: { style: BorderStyle.SINGLE, size: 6 }, left: { style: BorderStyle.SINGLE, size: 6 }, right: { style: BorderStyle.SINGLE, size: 6 } }, children: [new Paragraph({ text: p.responsavel || '' })] })
                    ]
                  })
                ))
              ]
            })
          ]
        }]
      });

      const buffer = await Packer.toBuffer(doc);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', 'attachment; filename="ATA_Reuniao.docx"');
      res.send(buffer);
    } catch (error) {
      res.status(500).json({ erro: error.message });
    }
  }
};
