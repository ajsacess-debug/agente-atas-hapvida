const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, BorderStyle, WidthType, AlignmentType } = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

const PORT = process.env.PORT || 3000;

function gerarAta(dados) {
  const doc = new Document({
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
          run: new TextRun({ size: 22, bold: true, color: '000000' })
        }),
        new Paragraph({ text: '' }),

        new Paragraph({
          text: `REALIZADA EM ${(dados.data || 'DATA').toUpperCase()}.`,
          alignment: AlignmentType.CENTER,
          run: new TextRun({ size: 22, bold: true, color: '000000' })
        }),
        new Paragraph({ text: '' }),
        new Paragraph({ text: '' }),

        new Paragraph({
          text: 'Objetivo da Reunião',
          run: new TextRun({ size: 24, bold: true, color: '000000' })
        }),
        new Paragraph({ text: dados.objetivo || 'Não informado', spacing: { line: 360 } }),
        new Paragraph({ text: '' }),

        new Paragraph({
          text: 'Resumo Executivo',
          run: new TextRun({ size: 24, bold: true, color: '000000' })
        }),
        new Paragraph({ text: dados.resumo || 'Não informado', spacing: { line: 360 } }),
        new Paragraph({ text: '' }),

        new Paragraph({
          text: 'Principais Definições e Entendimentos',
          run: new TextRun({ size: 24, bold: true, color: '000000' })
        }),

        ...(dados.definicoes || []).flatMap(def => [
          new Paragraph({
            text: def.topico || 'Tópico',
            run: new TextRun({ size: 22, bold: true, color: '000000' }),
            spacing: { before: 120, after: 80 }
          }),
          ...((def.pontos || []).map(ponto =>
            new Paragraph({ text: ponto, spacing: { line: 280 }, bullet: { level: 0 } })
          ))
        ]),

        new Paragraph({ text: '' }),

        new Paragraph({
          text: 'Riscos Identificados',
          run: new TextRun({ size: 24, bold: true, color: '000000' })
        }),

        ...(dados.riscos?.altoImpacto?.length > 0 ? [
          new Paragraph({
            text: 'Alto Impacto',
            run: new TextRun({ size: 22, bold: true, color: '000000' }),
            spacing: { before: 120, after: 80 }
          }),
          ...(dados.riscos.altoImpacto.map(risco =>
            new Paragraph({ text: risco, spacing: { line: 280 }, bullet: { level: 0 } })
          ))
        ] : []),

        ...(dados.riscos?.medioImpacto?.length > 0 ? [
          new Paragraph({
            text: 'Médio Impacto',
            run: new TextRun({ size: 22, bold: true, color: '000000' }),
            spacing: { before: 120, after: 80 }
          }),
          ...(dados.riscos.medioImpacto.map(risco =>
            new Paragraph({ text: risco, spacing: { line: 280 }, bullet: { level: 0 } })
          ))
        ] : []),

        new Paragraph({ text: '' }),

        new Paragraph({
          text: 'Próximos Passos',
          run: new TextRun({ size: 24, bold: true, color: '000000' })
        }),

        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({
              cells: [
                new TableCell({
                  width: { size: 70, type: WidthType.PERCENTAGE },
                  borders: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, left: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, right: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
                  shading: { type: 'clear', fill: 'D9D9D9' },
                  children: [new Paragraph({ text: 'Ação', run: new TextRun({ bold: true }) })]
                }),
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  borders: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, left: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, right: { style: BorderStyle.SINGLE, size: 6, color: '000000' } },
                  shading: { type: 'clear', fill: 'D9D9D9' },
                  children: [new Paragraph({ text: 'Responsável', run: new TextRun({ bold: true }) })]
                })
              ]
            }),
            ...((dados.proximosPassos || []).map(passo =>
              new TableRow({
                cells: [
                  new TableCell({ width: { size: 70, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, left: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, right: { style: BorderStyle.SINGLE, size: 6, color: '000000' } }, children: [new Paragraph({ text: passo.acao || '' })] }),
                  new TableCell({ width: { size: 30, type: WidthType.PERCENTAGE }, borders: { top: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, left: { style: BorderStyle.SINGLE, size: 6, color: '000000' }, right: { style: BorderStyle.SINGLE, size: 6, color: '000000' } }, children: [new Paragraph({ text: passo.responsavel || '' })] })
                ]
              })
            ))
          ]
        }),

        new Paragraph({ text: '' }),
        new Paragraph({ text: '' }),

        new Paragraph({
          text: 'Projetos e Transformação',
          alignment: AlignmentType.LEFT,
          run: new TextRun({ size: 20, italics: true, color: '7B8FB3' })
        }),

        new Paragraph({
          text: 'PÚBLICO',
          alignment: AlignmentType.CENTER,
          run: new TextRun({ size: 28, bold: true, color: 'E31D23' }),
          spacing: { before: 240 }
        })
      ]
    }]
  });

  return doc;
}

app.post('/gerar-ata', async (req, res) => {
  try {
    const dados = req.body;
    const doc = gerarAta(dados);
    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="ATA_Reuniao.docx"');
    res.send(buffer);
  } catch (error) {
    console.error('Erro:', error);
    res.status(500).json({ erro: 'Erro ao gerar documento', details: error.message });
  }
});

app.get('/', (req, res) => {
  res.json({ status: 'Agente de Atas funcionando!' });
});

app.get('/.well-known/ai-plugin.json', (req, res) => {
  res.json({
    schema_version: 'v1',
    name_for_human: 'Gerador de Atas Hapvida',
    name_for_model: 'atas_hapvida',
    description_for_human: 'Gera atas de reunião em Word no padrão Hapvida a partir de transcrições ou documentos',
    description_for_model: 'Processa dados de reuniões e gera documentos Word formatados no padrão Hapvida',
    auth: { type: 'none' },
    api: {
      type: 'openapi',
      url: `${process.env.BASE_URL || 'http://localhost:3000'}/openapi.json`
    },
    logo_url: `${process.env.BASE_URL || 'http://localhost:3000'}/logo.png`,
    contact_email: 'seu-email@example.com',
    legal_info_url: 'https://github.com/antonio-silva/agente-atas-hapvida'
  });
});

app.get('/openapi.json', (req, res) => {
  res.json({
    openapi: '3.0.0',
    info: {
      title: 'Gerador de Atas Hapvida',
      version: '1.0.0'
    },
    servers: [{ url: process.env.BASE_URL || 'http://localhost:3000' }],
    paths: {
      '/gerar-ata': {
        post: {
          operationId: 'gerarAta',
          summary: 'Gera uma ata de reunião em Word',
          requestBody: {
            required: true,
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  properties: {
                    assunto: { type: 'string', description: 'Assunto da reunião' },
                    data: { type: 'string', description: 'Data da reunião' },
                    objetivo: { type: 'string', description: 'Objetivo da reunião' },
                    resumo: { type: 'string', description: 'Resumo executivo' },
                    definicoes: {
                      type: 'array',
                      items: {
                        type: 'object',
                        properties: {
                          topico: { type: 'string' },
                          pontos: { type: 'array', items: { type: 'string' } }
                        }
                      }
                    },
                    riscos: {
                      type: 'object',
                      properties: {
                        altoImpacto: { type: 'array', items: { type: 'string' } },
                        medioImpacto: { type: 'array', items: { type: 'string' } }
                      }
                    },
                    proximosPassos: {
                      type: 'array',
                      items: {
                        type: 'object',
                        properties: {
                          acao: { type: 'string' },
                          responsavel: { type: 'string' }
                        }
                      }
                    }
                  },
                  required: ['assunto', 'data', 'objetivo', 'resumo']
                }
              }
            }
          },
          responses: {
            '200': { description: 'Documento Word gerado com sucesso' },
            '500': { description: 'Erro ao gerar documento' }
          }
        }
      }
    }
  });
});

app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
