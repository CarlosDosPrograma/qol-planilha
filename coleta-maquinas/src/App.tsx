import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

// ---------------------------------------------------------------------------
// CONFIGURAÇÕES DO SETOR E MÁQUINAS (DESMA)
// ---------------------------------------------------------------------------
interface MaquinaConfig {
  id: number;
  nome: string;
  totalPostos: number;
}

const CONFIG_MAQUINAS: MaquinaConfig[] = [
  { id: 1, nome: 'Máquina 01', totalPostos: 30 },
  { id: 2, nome: 'Máquina 02', totalPostos: 24 },
  { id: 3, nome: 'Máquina 03', totalPostos: 24 },
  { id: 4, nome: 'Máquina 04', totalPostos: 24 },
];

interface PostoData {
  numeracaoPosto: number;
  tamanho: number;
  tempoRealCompacta: number;
  tempoRealInjetada: number;
}

interface MaquinaData {
  idMaquina: number;
  nomeMaquina: string;
  postos: PostoData[];
}

interface ColetaDiaria {
  idColeta: number;
  horarioSalvamento?: string;
  maquinas: MaquinaData[];
}

interface BancoDadosHistorico {
  [data: string]: ColetaDiaria[];
}

const OPCOES_TAMANHO = Array.from({ length: 16 }, (_, i) => 33 + i);
const getTodayString = () => new Date().toISOString().split('T')[0];

const gerarEstruturaDia = (): ColetaDiaria[] => {
  const criarMaquinas = () =>
    CONFIG_MAQUINAS.map((m) => ({
      idMaquina: m.id,
      nomeMaquina: m.nome,
      postos: Array.from({ length: m.totalPostos }, (_, idx) => ({
        numeracaoPosto: idx + 1,
        tamanho: 33,
        tempoRealCompacta: 0,
        tempoRealInjetada: 0,
      })),
    }));

  return [1, 2, 3, 4].map((idColeta) => ({
    idColeta,
    maquinas: criarMaquinas(),
  }));
};

export default function App() {
  const dataHoje = getTodayString();
  const [dataSelecionada, setDataSelecionada] = useState<string>(dataHoje);
  
  // Controle de Módulo: 'injecao' | 'desmoldante'
  const [moduloAtivo, setModuloAtivo] = useState<'injecao' | 'desmoldante'>('injecao');
  const [coletaAtualIdx, setColetaAtualIdx] = useState<number>(1);
  const [maquinaAtivaId, setMaquinaAtivaId] = useState<number>(CONFIG_MAQUINAS[0].id);
  const [mensagem, setMensagem] = useState<string>('');

  // Persistência local
  const [bancoDados, setBancoDados] = useState<BancoDadosHistorico>(() => {
    const salvas = localStorage.getItem('banco_coletas_desma_clean_v6');
    if (salvas) return JSON.parse(salvas);
    return {};
  });

  useEffect(() => {
    localStorage.setItem('banco_coletas_desma_clean_v6', JSON.stringify(bancoDados));
  }, [bancoDados]);

  const coletasDoDia = bancoDados[dataSelecionada] || gerarEstruturaDia();
  const isReadOnly = dataSelecionada !== dataHoje;

  const handleCampoChange = (
    postoIdx: number,
    field: 'tempoRealCompacta' | 'tempoRealInjetada' | 'tamanho',
    value: string
  ) => {
    if (isReadOnly) return;

    let parsedVal: number = 0;
    if (field === 'tamanho') parsedVal = parseInt(value) || 33;
    if (field === 'tempoRealCompacta' || field === 'tempoRealInjetada') parsedVal = parseFloat(value) || 0;

    setBancoDados((prev) => {
      const bancoAtual = { ...prev };
      if (!bancoAtual[dataSelecionada]) {
        bancoAtual[dataSelecionada] = gerarEstruturaDia();
      }

      const dia = [...bancoAtual[dataSelecionada]];
      const coleta = dia[coletaAtualIdx - 1];
      const maquina = coleta.maquinas.find((m) => m.idMaquina === maquinaAtivaId);

      if (maquina) {
        maquina.postos[postoIdx][field] = parsedVal;
      }

      bancoAtual[dataSelecionada] = dia;
      return bancoAtual;
    });
  };

  const clonarColetaAnterior = () => {
    if (isReadOnly || coletaAtualIdx === 1) return;

    setBancoDados((prev) => {
      const bancoAtual = { ...prev };
      if (!bancoAtual[dataSelecionada]) return prev;

      const dia = [...bancoAtual[dataSelecionada]];
      const coletaAnterior = dia[coletaAtualIdx - 2];

      dia[coletaAtualIdx - 1].maquinas = JSON.parse(JSON.stringify(coletaAnterior.maquinas));
      bancoAtual[dataSelecionada] = dia;
      return bancoAtual;
    });
    exibirMensagem(`Dados clonados da ${coletaAtualIdx - 1}ª Coleta.`);
  };

  const salvarColetaAtual = () => {
    if (isReadOnly) return;
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });

    setBancoDados((prev) => {
      const bancoAtual = { ...prev };
      if (!bancoAtual[dataSelecionada]) {
        bancoAtual[dataSelecionada] = gerarEstruturaDia();
      }

      const dia = [...bancoAtual[dataSelecionada]];
      dia[coletaAtualIdx - 1].horarioSalvamento = agora;
      bancoAtual[dataSelecionada] = dia;
      return bancoAtual;
    });
    exibirMensagem(`Coleta ${coletaAtualIdx} registrada às ${agora}.`);
  };

  const exibirMensagem = (texto: string) => {
    setMensagem(texto);
    setTimeout(() => setMensagem(''), 3000);
  };

  const exportarParaExcel = () => {
    const linhasExcel: any[] = [];

    coletasDoDia.forEach((coleta) => {
      coleta.maquinas.forEach((maquina) => {
        maquina.postos.forEach((posto) => {
          linhasExcel.push({
            Data: dataSelecionada,
            Setor: 'Desma',
            Coleta: `${coleta.idColeta}ª Coleta`,
            'Horário Registro': coleta.horarioSalvamento || 'Pendente',
            Máquina: maquina.nomeMaquina,
            'Posto Nº': posto.numeracaoPosto,
            Tamanho: posto.tamanho,
            'T.R. Compacta': posto.tempoRealCompacta,
            'T.R. Injetada': posto.tempoRealInjetada,
          });
        });
      });
    });

    const worksheet = XLSX.utils.json_to_sheet(linhasExcel);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Injeção');
    XLSX.writeFile(workbook, `Injecao_Desma_${dataSelecionada}.xlsx`);
  };

  const coletaAtual = coletasDoDia[coletaAtualIdx - 1];
  const maquinaAtual =
    coletaAtual.maquinas.find((m) => m.idMaquina === maquinaAtivaId) || coletaAtual.maquinas[0];

  return (
    <div style={{ backgroundColor: '#F8FAFC', minHeight: '100vh', fontFamily: 'system-ui, -apple-system, sans-serif', color: '#0F172A' }}>
      
      {/* BARRA SUPERIOR INDUSTRIAL CLEAN */}
      <header style={{ backgroundColor: '#0F172A', color: '#FFFFFF', padding: '12px 24px', borderBottom: '3px solid #DC2626', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '11px', fontWeight: 700, backgroundColor: '#DC2626', color: '#FFF', padding: '2px 6px', borderRadius: '4px', letterSpacing: '0.5px' }}>SETOR DESMA</span>
            <h1 style={{ margin: 0, fontSize: '18px', fontWeight: 600, letterSpacing: '-0.3px' }}>Sistema de Coleta Operacional</h1>
          </div>
        </div>

        {/* SELETOR DE DATA / HISTÓRICO */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#1E293B', padding: '4px 10px', borderRadius: '6px', border: '1px solid #334155' }}>
          <span style={{ fontSize: '12px', color: '#94A3B8', fontWeight: 500 }}>Data do Registro:</span>
          <input
            type="date"
            value={dataSelecionada}
            onChange={(e) => setDataSelecionada(e.target.value)}
            style={{
              padding: '4px 8px',
              borderRadius: '4px',
              border: '1px solid #475569',
              backgroundColor: '#0F172A',
              color: '#FFFFFF',
              WebkitTextFillColor: '#FFFFFF',
              fontSize: '13px',
              fontWeight: 600,
            }}
          />
        </div>
      </header>

      <main style={{ maxWidth: '1000px', margin: '0 auto', padding: '20px' }}>
        
        {/* AVISO DE MODO HISTÓRICO */}
        {isReadOnly && (
          <div style={{ backgroundColor: '#FEF2F2', color: '#991B1B', border: '1px solid #FCA5A5', padding: '10px 14px', borderRadius: '6px', marginBottom: '16px', fontSize: '13px', fontWeight: 500 }}>
            <strong>Modo Consulta (Histórico):</strong> Exibindo dados salvos do dia {dataSelecionada}. Apenas leitura.
          </div>
        )}

        {/* NAVEGAÇÃO DE MÓDULOS (ABAS SUPERIORES) */}
        <div style={{ display: 'flex', gap: '8px', marginBottom: '20px', borderBottom: '2px solid #E2E8F0', paddingBottom: '2px' }}>
          <button
            onClick={() => setModuloAtivo('injecao')}
            style={{
              padding: '8px 16px',
              fontSize: '14px',
              fontWeight: 600,
              borderRadius: '6px 6px 0 0',
              border: 'none',
              cursor: 'pointer',
              backgroundColor: moduloAtivo === 'injecao' ? '#0F172A' : 'transparent',
              color: moduloAtivo === 'injecao' ? '#FFFFFF' : '#64748B',
              transition: 'all 0.15s ease',
            }}
          >
            Tabela de Injeção
          </button>
          
          <button
            onClick={() => setModuloAtivo('desmoldante')}
            style={{
              padding: '8px 16px',
              fontSize: '14px',
              fontWeight: 600,
              borderRadius: '6px 6px 0 0',
              border: 'none',
              cursor: 'pointer',
              backgroundColor: moduloAtivo === 'desmoldante' ? '#0F172A' : 'transparent',
              color: moduloAtivo === 'desmoldante' ? '#FFFFFF' : '#64748B',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
            }}
          >
            Desmoldante
            <span style={{ fontSize: '10px', backgroundColor: '#E2E8F0', color: '#475569', padding: '2px 6px', borderRadius: '10px', fontWeight: 500 }}>Em Breve</span>
          </button>
        </div>

        {/* =================================================================== */}
        {/* CONTEÚDO 1: TABELA DE INJEÇÃO (MÓDULO ATIVO) */}
        {/* =================================================================== */}
        {moduloAtivo === 'injecao' && (
          <div>
            {/* SUB-NAVEGAÇÃO: CICLOS DE COLETA (1ª a 4ª) */}
            <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
              {[1, 2, 3, 4].map((num) => {
                const finalizada = Boolean(coletasDoDia[num - 1]?.horarioSalvamento);
                const selecionada = coletaAtualIdx === num;

                return (
                  <button
                    key={num}
                    onClick={() => setColetaAtualIdx(num)}
                    style={{
                      flex: 1,
                      padding: '10px 8px',
                      fontSize: '13px',
                      fontWeight: 600,
                      borderRadius: '6px',
                      border: selecionada ? '2px solid #0F172A' : '1px solid #CBD5E1',
                      cursor: 'pointer',
                      backgroundColor: selecionada ? '#0F172A' : '#FFFFFF',
                      color: selecionada ? '#FFFFFF' : '#334155',
                      display: 'flex',
                      justifyContent: 'center',
                      alignItems: 'center',
                      gap: '6px',
                    }}
                  >
                    <span>{num}ª Coleta</span>
                    {finalizada && <span style={{ color: selecionada ? '#4ADE80' : '#16A34A', fontSize: '12px' }}>✓</span>}
                  </button>
                );
              })}
            </div>

            {/* SUB-NAVEGAÇÃO: SELETOR DE MÁQUINAS */}
            <div style={{ display: 'flex', gap: '8px', marginBottom: '20px' }}>
              {CONFIG_MAQUINAS.map((m) => {
                const ativa = maquinaAtivaId === m.id;
                return (
                  <button
                    key={m.id}
                    onClick={() => setMaquinaAtivaId(m.id)}
                    style={{
                      flex: 1,
                      padding: '8px',
                      fontSize: '13px',
                      fontWeight: 600,
                      borderRadius: '6px',
                      border: '1px solid #CBD5E1',
                      cursor: 'pointer',
                      backgroundColor: ativa ? '#FFFFFF' : '#F1F5F9',
                      color: ativa ? '#DC2626' : '#475569',
                      borderBottom: ativa ? '3px solid #DC2626' : '1px solid #CBD5E1',
                    }}
                  >
                    {m.nome} <span style={{ fontSize: '11px', color: '#94A3B8', fontWeight: 400 }}>({m.totalPostos} postos)</span>
                  </button>
                );
              })}
            </div>

            {/* BARRA DE AÇÕES */}
            <div style={{ display: 'flex', gap: '10px', marginBottom: '16px', flexWrap: 'wrap' }}>
              {coletaAtualIdx > 1 && (
                <button
                  onClick={clonarColetaAnterior}
                  disabled={isReadOnly}
                  style={{
                    padding: '9px 16px',
                    backgroundColor: isReadOnly ? '#E2E8F0' : '#FFFFFF',
                    color: isReadOnly ? '#94A3B8' : '#334155',
                    border: '1px solid #CBD5E1',
                    borderRadius: '6px',
                    fontWeight: 600,
                    fontSize: '13px',
                    cursor: isReadOnly ? 'not-allowed' : 'pointer',
                  }}
                >
                  Copiar Coleta Anterior
                </button>
              )}
              <button
                onClick={salvarColetaAtual}
                disabled={isReadOnly}
                style={{
                  padding: '9px 18px',
                  backgroundColor: isReadOnly ? '#CBD5E1' : '#DC2626',
                  color: '#FFFFFF',
                  border: 'none',
                  borderRadius: '6px',
                  fontWeight: 600,
                  fontSize: '13px',
                  cursor: isReadOnly ? 'not-allowed' : 'pointer',
                }}
              >
                Salvar Coleta {coletaAtualIdx}
              </button>
              <button
                onClick={exportarParaExcel}
                style={{
                  padding: '9px 16px',
                  backgroundColor: '#0F172A',
                  color: '#FFFFFF',
                  border: 'none',
                  borderRadius: '6px',
                  fontWeight: 600,
                  fontSize: '13px',
                  cursor: 'pointer',
                  marginLeft: 'auto',
                }}
              >
                Exportar Excel (.xlsx)
              </button>
            </div>

            {/* MENSAGEM DE FEEDBACK */}
            {mensagem && (
              <div style={{ padding: '8px 12px', backgroundColor: '#F0FDF4', color: '#166534', border: '1px solid #BBF7D0', borderRadius: '6px', marginBottom: '16px', fontSize: '13px', fontWeight: 500 }}>
                {mensagem}
              </div>
            )}

            {/* TABELA PRINCIPAL DE DADOS DE INJEÇÃO */}
            <div style={{ backgroundColor: '#FFFFFF', borderRadius: '8px', border: '1px solid #E2E8F0', overflow: 'hidden' }}>
              <div style={{ padding: '12px 16px', backgroundColor: '#F8FAFC', borderBottom: '1px solid #E2E8F0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontWeight: 600, fontSize: '14px', color: '#0F172A' }}>
                  {maquinaAtual.nomeMaquina} — Tabela de Injeção
                </span>
                <span style={{ fontSize: '12px', color: '#64748B' }}>
                  {maquinaAtual.postos.length} postos em exibição
                </span>
              </div>

              <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                <thead>
                  <tr style={{ backgroundColor: '#F1F5F9', color: '#475569', borderBottom: '1px solid #E2E8F0', textTransform: 'uppercase', fontSize: '11px', letterSpacing: '0.5px' }}>
                    <th style={{ padding: '10px 16px', width: '20%' }}>Posto</th>
                    <th style={{ padding: '10px 16px', width: '25%' }}>Tamanho</th>
                    <th style={{ padding: '10px 16px', width: '27%' }}>Tempo Real Compacta</th>
                    <th style={{ padding: '10px 16px', width: '28%' }}>Tempo Real Injetada</th>
                  </tr>
                </thead>
                <tbody>
                  {maquinaAtual.postos.map((posto, idx) => (
                    <tr
                      key={posto.numeracaoPosto}
                      style={{
                        borderBottom: '1px solid #F1F5F9',
                        backgroundColor: idx % 2 === 0 ? '#FFFFFF' : '#FAFAFA',
                      }}
                    >
                      <td style={{ padding: '10px 16px', fontWeight: 600, color: '#0F172A' }}>
                        Posto {posto.numeracaoPosto}
                      </td>

                      {/* Tamanho */}
                      <td style={{ padding: '8px 16px' }}>
                        <select
                          value={posto.tamanho}
                          disabled={isReadOnly}
                          onChange={(e) => handleCampoChange(idx, 'tamanho', e.target.value)}
                          style={{
                            padding: '6px 10px',
                            fontSize: '13px',
                            fontWeight: 600,
                            borderRadius: '4px',
                            border: '1px solid #CBD5E1',
                            backgroundColor: isReadOnly ? '#F1F5F9' : '#FFFFFF',
                            color: '#0F172A',
                            WebkitTextFillColor: '#0F172A',
                            cursor: isReadOnly ? 'not-allowed' : 'pointer',
                          }}
                        >
                          {OPCOES_TAMANHO.map((tam) => (
                            <option key={tam} value={tam} style={{ color: '#0F172A', backgroundColor: '#FFFFFF' }}>
                              Nº {tam}
                            </option>
                          ))}
                        </select>
                      </td>

                      {/* Tempo Real Compacta */}
                      <td style={{ padding: '8px 16px' }}>
                        <input
                          type="number"
                          step="0.01"
                          disabled={isReadOnly}
                          value={posto.tempoRealCompacta || ''}
                          onChange={(e) => handleCampoChange(idx, 'tempoRealCompacta', e.target.value)}
                          placeholder="0.00"
                          style={{
                            width: '110px',
                            padding: '6px 10px',
                            fontSize: '13px',
                            borderRadius: '4px',
                            border: '1px solid #CBD5E1',
                            backgroundColor: isReadOnly ? '#F1F5F9' : '#FFFFFF',
                            color: '#0F172A',
                            WebkitTextFillColor: '#0F172A',
                            fontWeight: 600,
                          }}
                        />
                      </td>

                      {/* Tempo Real Injetada */}
                      <td style={{ padding: '8px 16px' }}>
                        <input
                          type="number"
                          step="0.01"
                          disabled={isReadOnly}
                          value={posto.tempoRealInjetada || ''}
                          onChange={(e) => handleCampoChange(idx, 'tempoRealInjetada', e.target.value)}
                          placeholder="0.00"
                          style={{
                            width: '110px',
                            padding: '6px 10px',
                            fontSize: '13px',
                            borderRadius: '4px',
                            border: '1px solid #CBD5E1',
                            backgroundColor: isReadOnly ? '#F1F5F9' : '#FFFFFF',
                            color: '#0F172A',
                            WebkitTextFillColor: '#0F172A',
                            fontWeight: 600,
                          }}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* =================================================================== */}
        {/* CONTEÚDO 2: DESMOLDANTE (ESPAÇO EM DESENVOLVIMENTO) */}
        {/* =================================================================== */}
        {moduloAtivo === 'desmoldante' && (
          <div style={{ backgroundColor: '#FFFFFF', borderRadius: '8px', border: '1px solid #E2E8F0', padding: '60px 20px', textAlign: 'center' }}>
            <div style={{ maxWidth: '400px', margin: '0 auto' }}>
              <div style={{ fontSize: '32px', marginBottom: '12px' }}>🧪</div>
              <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', fontWeight: 600, color: '#0F172A' }}>
                Planilha de Desmoldante
              </h3>
              <p style={{ margin: '0 0 20px 0', fontSize: '14px', color: '#64748B', lineHeight: '1.5' }}>
                Este espaço foi reservado para a futura coleta e controle da aplicação de desmoldante por máquina.
              </p>
              <span style={{ display: 'inline-block', backgroundColor: '#F1F5F9', color: '#475569', fontSize: '12px', fontWeight: 600, padding: '6px 14px', borderRadius: '20px', border: '1px solid #CBD5E1' }}>
                Em Desenvolvimento
              </span>
            </div>
          </div>
        )}

      </main>
    </div>
  );
}