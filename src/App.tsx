import React, { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  Upload,
  Info,
  Trash2,
  BarChart3,
  Users,
  AlertTriangle,
  Search,
  FileText,
  Brain,
  CheckCircle2,
  Clock,
  RefreshCcw,
  Star,
  Download,
  Filter,
  Eye,
  ShieldCheck,
  TrendingUp,
} from 'lucide-react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  LineChart,
  Line,
  PieChart,
  Pie,
} from 'recharts';
import './App.css';

type Atendimento = {
  protocolo: string;
  numeroCliente: string;
  cliente?: string;
  data: string;
  dataObj?: Date | null;
  canal?: string;
  pontosBot: string;
  pontosLista: string[];
  motivo: string;
  solucionado: boolean | null;
  erro: string;
  transferido: boolean;
  inatividade: boolean;
  pesquisaEncaminhada: boolean;
  tempo: string;
  agente?: string;
  nomeAgente?: string;
  totalMensagens?: number;
  recorrente: boolean;
  retornou24h: boolean;
  retornou72h: boolean;
  protocoloAnterior: string;
  proximoProtocolo: string;
  tempoAteProximoContato: string;
  classificacaoRecorrencia: string;
  solucaoConfirmada: string;
  nota: number | null;
  nps: string;
  csat: number | null;
  comentario: string;
  qualidade: string;
  arquivoOrigem: string[];
};

type ImportedFileData = {
  id: string;
  fileName: string;
  rows: any[][];
};

type ImportLog = {
  id: string;
  arquivo: string;
  tipo: string;
  registros: number;
};

const emptyAtendimento = (protocolo: string): Atendimento => ({
  protocolo,
  numeroCliente: '',
  cliente: '',
  data: '',
  dataObj: null,
  canal: '',
  pontosBot: '',
  pontosLista: [],
  motivo: '',
  solucionado: null,
  erro: '',
  transferido: false,
  inatividade: false,
  pesquisaEncaminhada: false,
  tempo: '',
  agente: '',
  nomeAgente: '',
  totalMensagens: 0,
  recorrente: false,
  retornou24h: false,
  retornou72h: false,
  protocoloAnterior: '-',
  proximoProtocolo: '-',
  tempoAteProximoContato: '-',
  classificacaoRecorrencia: 'indefinido',
  solucaoConfirmada: 'indefinido',
  nota: null,
  nps: '',
  csat: null,
  comentario: '',
  qualidade: 'Sem avaliação',
  arquivoOrigem: [],
});

function clean(value: unknown): string {
  return String(value ?? '').trim();
}

function normalizeHeader(value: unknown): string {
  return clean(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function parseDateBR(value: string): Date | null {
  if (!value) return null;

  const normalized = value.replace(' - ', ' ').trim();
  const match = normalized.match(
    /(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?/
  );

  if (!match) return null;

  const [, dd, mm, yyyy, hh = '00', min = '00', ss = '00'] = match;

  return new Date(
    Number(yyyy),
    Number(mm) - 1,
    Number(dd),
    Number(hh),
    Number(min),
    Number(ss)
  );
}

function formatDate(date: Date | null | undefined): string {
  if (!date) return '';

  return date.toLocaleString('pt-BR', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

function diffLabel(a?: Date | null, b?: Date | null): string {
  if (!a || !b) return '-';

  const diffMs = b.getTime() - a.getTime();
  const diffHours = Math.round(diffMs / 1000 / 60 / 60);

  if (diffHours < 1) return 'menos de 1h';
  if (diffHours < 24) return `${diffHours}h`;

  const days = Math.round(diffHours / 24);
  return `${days} dia(s)`;
}

function normalizeProtocolo(value: unknown): string {
  return clean(value).replace(/\D/g, '');
}

function detectReportType(
  rows: any[][]
): 'pontos' | 'analitico' | 'satisfacao' | 'desconhecido' {
  const text = rows.slice(0, 5).flat().map(normalizeHeader).join(' | ');

  if (text.includes('pontos de bot')) return 'pontos';
  if (
    text.includes('tipo de atendimento') &&
    text.includes('tempo atendimento')
  )
    return 'analitico';
  if (
    text.includes('pesquisa de satisfacao') ||
    (text.includes('questao') && text.includes('nota'))
  ) {
    return 'satisfacao';
  }

  return 'desconhecido';
}

function findHeaderRow(rows: any[][]): number {
  return rows.findIndex((row) =>
    row.some((cell) => normalizeHeader(cell) === 'protocolo')
  );
}

function rowToObject(headers: string[], row: any[]): Record<string, string> {
  const obj: Record<string, string> = {};

  headers.forEach((header, index) => {
    obj[normalizeHeader(header)] = clean(row[index]);
  });

  return obj;
}

function getField(
  row: Record<string, string>,
  possibleNames: string[]
): string {
  for (const name of possibleNames) {
    const key = normalizeHeader(name);
    if (row[key]) return row[key];
  }

  return '';
}

function getOrCreate(
  map: Map<string, Atendimento>,
  protocolo: string
): Atendimento {
  if (!map.has(protocolo)) {
    map.set(protocolo, emptyAtendimento(protocolo));
  }

  return map.get(protocolo)!;
}

function classifyNps(nota: number | null): string {
  if (nota === null) return '';

  if (nota >= 4) return 'Promotor';
  if (nota === 3) return 'Neutro';
  return 'Detrator';
}

function classifyQuality(a: Atendimento): string {
  const notaAlta = a.nota !== null && a.nota >= 4;
  const notaBaixa = a.nota !== null && a.nota <= 3;

  if (a.solucionado === true && a.retornou72h) return 'Possível falso positivo';
  if (a.solucionado === true && notaBaixa) return 'Solução questionável';
  if (a.solucionado === true && !a.retornou72h && notaAlta)
    return 'Solução validada';
  if (a.solucionado === false && notaBaixa) return 'Problema confirmado';
  if (a.nota === null) return 'Sem avaliação';

  return 'Indefinido';
}

function classifySolution(a: Atendimento): string {
  if (a.solucionado === true && !a.retornou72h) return 'sim';
  if (a.solucionado === true && a.retornou72h) return 'possível falso positivo';
  if (a.solucionado === false && a.retornou72h) return 'não';
  return 'indefinido';
}

async function readFileRows(file: File): Promise<any[][]> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', raw: false });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(firstSheet, {
    header: 1,
    defval: '',
    blankrows: false,
  }) as any[][];
}

function processFilesData(filesData: ImportedFileData[]) {
  const map = new Map<string, Atendimento>();
  const logs: ImportLog[] = [];

  for (const fileData of filesData) {
    const { fileName, rows } = fileData;
    const type = detectReportType(rows);
    const headerIndex = findHeaderRow(rows);

    if (headerIndex === -1) {
      logs.push({
        id: fileData.id,
        arquivo: fileName,
        tipo: 'não identificado',
        registros: 0,
      });
      continue;
    }

    const headers = rows[headerIndex].map(clean);
    const dataRows = rows.slice(headerIndex + 1);

    if (type === 'pontos') {
      let count = 0;

      for (const raw of dataRows) {
        const row = rowToObject(headers, raw);
        const protocolo = normalizeProtocolo(getField(row, ['Protocolo']));

        if (!protocolo) continue;

        const item = getOrCreate(map, protocolo);

        const data = getField(row, ['Data']);
        const dataObj = parseDateBR(data);
        const ponto = getField(row, ['Pontos de Bot']);
        const contato = getField(row, ['Contato']);
        const cliente = getField(row, ['Cliente']);

        if (!item.dataObj || (dataObj && dataObj < item.dataObj)) {
          item.dataObj = dataObj;
          item.data = formatDate(dataObj);
        }

        if (!item.numeroCliente && contato) item.numeroCliente = contato;
        if (!item.cliente && cliente) item.cliente = cliente;
        if (!item.canal) item.canal = getField(row, ['Tipo de Entrada']);
        if (!item.agente) item.agente = getField(row, ['Agente']);

        if (ponto && !item.pontosLista.includes(ponto)) {
          item.pontosLista.push(ponto);
        }

        const pontoNorm = normalizeHeader(ponto);

        if (pontoNorm.includes('transferencia fila')) item.transferido = true;
        if (pontoNorm.includes('inatividade')) item.inatividade = true;
        if (pontoNorm.includes('pesquisa encaminhada'))
          item.pesquisaEncaminhada = true;

        if (pontoNorm.includes('problema resolvido')) item.solucionado = true;
        if (pontoNorm.includes('problema nao resolvido'))
          item.solucionado = false;

        if (!item.arquivoOrigem.includes(fileName))
          item.arquivoOrigem.push(fileName);

        count++;
      }

      logs.push({
        id: fileData.id,
        arquivo: fileName,
        tipo: 'Pontos de Acesso',
        registros: count,
      });
    }

    if (type === 'analitico') {
      let count = 0;

      for (const raw of dataRows) {
        const row = rowToObject(headers, raw);
        const protocolo = normalizeProtocolo(getField(row, ['Protocolo']));

        if (!protocolo) continue;

        const item = getOrCreate(map, protocolo);

        const inicio = getField(row, [
          'Início Atendimento',
          'Inicio Atendimento',
        ]);
        const dataObj = parseDateBR(inicio);

        if (!item.dataObj || (dataObj && dataObj < item.dataObj)) {
          item.dataObj = dataObj;
          item.data = formatDate(dataObj);
        }

        item.canal = item.canal || getField(row, ['Canal']);
        item.tempo = item.tempo || getField(row, ['Tempo Atendimento']);
        item.numeroCliente = item.numeroCliente || getField(row, ['Origem']);
        item.agente = item.agente || getField(row, ['Agente']);
        item.nomeAgente = item.nomeAgente || getField(row, ['Nome Agente']);
        item.motivo = item.motivo || getField(row, ['Tabulação', 'Tabulacao']);
        item.erro = item.erro || getField(row, ['Substatus']);

        const totalMensagens = Number(getField(row, ['Total de mensagens']));
        if (!Number.isNaN(totalMensagens)) item.totalMensagens = totalMensagens;

        if (!item.arquivoOrigem.includes(fileName))
          item.arquivoOrigem.push(fileName);

        count++;
      }

      logs.push({
        id: fileData.id,
        arquivo: fileName,
        tipo: 'Protocolos Analítico',
        registros: count,
      });
    }

    if (type === 'satisfacao') {
      let count = 0;

      for (const raw of dataRows) {
        const row = rowToObject(headers, raw);
        const protocolo = normalizeProtocolo(getField(row, ['Protocolo']));

        if (!protocolo) continue;

        const item = getOrCreate(map, protocolo);

        const data = getField(row, ['Data/Hora']);
        const dataObj = parseDateBR(data);
        const questao = getField(row, ['Questão', 'Questao']);
        let resposta = getField(row, ['Nota']);

        const notaIndex = headers.findIndex(
          (h) => normalizeHeader(h) === 'nota'
        );

        if (notaIndex >= 0 && raw.length > notaIndex + 1) {
          resposta = raw.slice(notaIndex).map(clean).filter(Boolean).join(', ');
        }

        if (!item.numeroCliente)
          item.numeroCliente = getField(row, ['Contato']);
        if (!item.dataObj && dataObj) {
          item.dataObj = dataObj;
          item.data = formatDate(dataObj);
        }

        const questaoNorm = normalizeHeader(questao);

        if (questaoNorm.includes('avalia')) {
          const nota = Number(String(resposta).replace(',', '.'));
          if (!Number.isNaN(nota)) {
            item.nota = nota;
            item.nps = classifyNps(nota);
            item.csat = Math.round((nota / 5) * 100);
          }
        } else if (
          questaoNorm.includes('comentario') ||
          questaoNorm.includes('experiencia') ||
          questaoNorm.includes('compartilhar')
        ) {
          item.comentario = resposta;
        }

        if (!item.arquivoOrigem.includes(fileName))
          item.arquivoOrigem.push(fileName);

        count++;
      }

      logs.push({
        id: fileData.id,
        arquivo: fileName,
        tipo: 'Pesquisa de Satisfação',
        registros: count,
      });
    }

    if (type === 'desconhecido') {
      logs.push({
        id: fileData.id,
        arquivo: fileName,
        tipo: 'desconhecido',
        registros: dataRows.length,
      });
    }
  }

  const atendimentos = Array.from(map.values()).map((item) => ({
    ...item,
    pontosBot: item.pontosLista.join(' | '),
    motivo: item.motivo || inferMotivo(item.pontosLista),
    erro: item.erro || inferErro(item.pontosLista),
  }));

  const withRecurrence = calculateRecurrence(atendimentos);

  return {
    atendimentos: withRecurrence,
    logs,
  };
}

function inferMotivo(pontos: string[]): string {
  const joined = normalizeHeader(pontos.join(' '));

  if (joined.includes('safeid')) return 'SafeID';
  if (joined.includes('a1')) return 'Certificado A1';
  if (joined.includes('a3')) return 'Certificado A3';
  if (joined.includes('revogacao')) return 'Revogação';
  if (joined.includes('senha')) return 'Senha';
  if (joined.includes('instalacao')) return 'Instalação';
  if (joined.includes('transferencia')) return 'Transferência';

  return 'Não identificado';
}

function inferErro(pontos: string[]): string {
  const joined = normalizeHeader(pontos.join(' '));

  if (joined.includes('problema nao resolvido'))
    return 'Problema não resolvido';
  if (joined.includes('inatividade')) return 'Finalizado por inatividade';
  if (joined.includes('transferencia')) return 'Transferência para fila';

  return '';
}

function calculateRecurrence(atendimentos: Atendimento[]): Atendimento[] {
  const byClient = new Map<string, Atendimento[]>();

  for (const item of atendimentos) {
    const key = item.numeroCliente || item.cliente || '';
    if (!key) continue;

    if (!byClient.has(key)) byClient.set(key, []);
    byClient.get(key)!.push(item);
  }

  for (const [, list] of byClient) {
    list.sort((a, b) => {
      const dateA = a.dataObj?.getTime() ?? 0;
      const dateB = b.dataObj?.getTime() ?? 0;
      return dateA - dateB;
    });

    const isRecurring = list.length > 1;

    for (let i = 0; i < list.length; i++) {
      const current = list[i];
      const previous = list[i - 1];
      const next = list[i + 1];

      current.recorrente = isRecurring;
      current.protocoloAnterior = previous?.protocolo || '-';
      current.proximoProtocolo = next?.protocolo || '-';
      current.tempoAteProximoContato = diffLabel(
        current.dataObj,
        next?.dataObj
      );

      if (next?.dataObj && current.dataObj) {
        const diffHours =
          (next.dataObj.getTime() - current.dataObj.getTime()) / 1000 / 60 / 60;

        current.retornou24h = diffHours <= 24;
        current.retornou72h = diffHours <= 72;

        if (diffHours <= 24)
          current.classificacaoRecorrencia = 'retorno rápido';
        else if (diffHours <= 72)
          current.classificacaoRecorrencia = 'retorno recente';
        else current.classificacaoRecorrencia = 'novo contato posterior';
      } else {
        current.retornou24h = false;
        current.retornou72h = false;
        current.classificacaoRecorrencia = isRecurring
          ? 'último contato da jornada'
          : 'sem retorno identificado';
      }

      current.solucaoConfirmada = classifySolution(current);
      current.qualidade = classifyQuality(current);
    }
  }

  return atendimentos.sort((a, b) => {
    const dateA = a.dataObj?.getTime() ?? 0;
    const dateB = b.dataObj?.getTime() ?? 0;
    return dateB - dateA;
  });
}

function Badge({
  children,
  tone = 'gray',
}: {
  children: React.ReactNode;
  tone?: 'green' | 'red' | 'orange' | 'yellow' | 'blue' | 'gray';
}) {
  return <span className={`badge ${tone}`}>{children}</span>;
}

function InfoTip({ text }: { text: string }) {
  return (
    <span className="info-tip" title={text}>
      <Info size={14} />
    </span>
  );
}

function StatCard({
  icon: Icon,
  title,
  value,
  helper,
  tone = 'neutral',
  info,
}: {
  icon: React.ElementType;
  title: string;
  value: string | number;
  helper?: string;
  tone?: string;
  info?: string;
}) {
  return (
    <div className="card stat-card">
      <div>
        <div className="stat-title-row">
          <p className="stat-title">{title}</p>
          {info && <InfoTip text={info} />}
        </div>

        <p className="stat-value">{value}</p>

        {helper && <p className="stat-helper">{helper}</p>}
      </div>

      <div className={`stat-icon ${tone}`}>
        <Icon size={22} />
      </div>
    </div>
  );
}

function SectionTitle({
  icon: Icon,
  title,
  description,
}: {
  icon: React.ElementType;
  title: string;
  description?: string;
}) {
  return (
    <div className="section-title">
      <div className="section-icon">
        <Icon size={20} />
      </div>
      <div>
        <h2>{title}</h2>
        {description && <p>{description}</p>}
      </div>
    </div>
  );
}

function statusBadge(item: Atendimento) {
  if (item.solucionado === true) return <Badge tone="green">Solucionado</Badge>;
  if (item.solucionado === false)
    return <Badge tone="red">Não solucionado</Badge>;
  return <Badge tone="yellow">Indefinido</Badge>;
}

function qualityTone(
  qualidade: string
): 'green' | 'red' | 'orange' | 'yellow' | 'gray' {
  const q = normalizeHeader(qualidade);

  if (q.includes('validada')) return 'green';
  if (q.includes('questionavel')) return 'orange';
  if (q.includes('falso') || q.includes('problema')) return 'red';
  if (q.includes('sem avaliacao')) return 'gray';

  return 'yellow';
}

export default function App() {
  const [page, setPage] = useState('dashboard');
  const [search, setSearch] = useState('');
  const [atendimentos, setAtendimentos] = useState<Atendimento[]>([]);
  const [logs, setLogs] = useState<ImportLog[]>([]);
  const [importedFiles, setImportedFiles] = useState<ImportedFileData[]>([]);
  const [isImporting, setIsImporting] = useState(false);

  async function handleFiles(event: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files || []);

    if (!files.length) return;

    setIsImporting(true);

    try {
      const newFilesData: ImportedFileData[] = await Promise.all(
        files.map(async (file) => ({
          id: `${file.name}-${Date.now()}-${Math.random()}`,
          fileName: file.name,
          rows: await readFileRows(file),
        }))
      );

      const updatedFiles = [...importedFiles, ...newFilesData];

      const result = processFilesData(updatedFiles);

      setImportedFiles(updatedFiles);
      setAtendimentos(result.atendimentos);
      setLogs(result.logs);
      setPage('dashboard');
    } catch (error) {
      console.error(error);
      alert('Erro ao importar os arquivos. Veja o console para mais detalhes.');
    } finally {
      setIsImporting(false);
      event.target.value = '';
    }
  }

  function handleDeleteImportedFile(fileId: string) {
    const confirmDelete = window.confirm(
      'Deseja realmente excluir este arquivo importado? A dashboard será recalculada com os arquivos restantes.'
    );

    if (!confirmDelete) return;

    const updatedFiles = importedFiles.filter((file) => file.id !== fileId);

    if (updatedFiles.length === 0) {
      setImportedFiles([]);
      setAtendimentos([]);
      setLogs([]);
      return;
    }

    const result = processFilesData(updatedFiles);

    setImportedFiles(updatedFiles);
    setAtendimentos(result.atendimentos);
    setLogs(result.logs);
  }

  const filtered = useMemo(() => {
    return atendimentos.filter((item) => {
      const q = search.toLowerCase();

      return (
        item.protocolo.toLowerCase().includes(q) ||
        item.numeroCliente.toLowerCase().includes(q) ||
        item.cliente?.toLowerCase().includes(q) ||
        item.motivo.toLowerCase().includes(q) ||
        item.erro.toLowerCase().includes(q) ||
        item.pontosBot.toLowerCase().includes(q)
      );
    });
  }, [search, atendimentos]);

  const total = atendimentos.length;
  const solucionados = atendimentos.filter(
    (a) => a.solucionado === true
  ).length;
  const naoSolucionados = atendimentos.filter(
    (a) => a.solucionado === false
  ).length;
  const transferidos = atendimentos.filter((a) => a.transferido).length;
  const recorrentes = new Set(
    atendimentos.filter((a) => a.recorrente).map((a) => a.numeroCliente)
  ).size;
  const retornos72h = atendimentos.filter((a) => a.retornou72h).length;
  const falsoPositivo = atendimentos.filter((a) =>
    normalizeHeader(a.qualidade).includes('falso')
  ).length;

  const avaliacoes = atendimentos.filter((a) => a.nota !== null);
  const notaMedia =
    avaliacoes.length > 0
      ? (
          avaliacoes.reduce((acc, a) => acc + Number(a.nota), 0) /
          avaliacoes.length
        ).toFixed(1)
      : '-';

  const csatMedio =
    avaliacoes.length > 0
      ? `${Math.round(
          avaliacoes.reduce((acc, a) => acc + Number(a.csat || 0), 0) /
            avaliacoes.length
        )}%`
      : '-';

  const chartDia = useMemo(() => {
    const map = new Map<
      string,
      {
        dia: string;
        atendimentos: number;
        recorrentes: number;
        satisfacaoSoma: number;
        satisfacaoQtd: number;
        satisfacao: number;
      }
    >();

    for (const item of atendimentos) {
      const day = item.dataObj
        ? item.dataObj.toLocaleDateString('pt-BR', {
            day: '2-digit',
            month: '2-digit',
          })
        : 'Sem data';

      if (!map.has(day)) {
        map.set(day, {
          dia: day,
          atendimentos: 0,
          recorrentes: 0,
          satisfacaoSoma: 0,
          satisfacaoQtd: 0,
          satisfacao: 0,
        });
      }

      const group = map.get(day)!;
      group.atendimentos++;
      if (item.recorrente) group.recorrentes++;
      if (item.nota !== null) {
        group.satisfacaoSoma += item.nota;
        group.satisfacaoQtd++;
      }
    }

    return Array.from(map.values()).map((item) => ({
      ...item,
      satisfacao:
        item.satisfacaoQtd > 0
          ? Number((item.satisfacaoSoma / item.satisfacaoQtd).toFixed(1))
          : 0,
    }));
  }, [atendimentos]);

  const motivos = useMemo(() => {
    const map = new Map<string, number>();

    for (const item of atendimentos) {
      const key = item.motivo || 'Não identificado';
      map.set(key, (map.get(key) || 0) + 1);
    }

    return Array.from(map.entries())
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8);
  }, [atendimentos]);

  const qualidadeData = useMemo(() => {
    const map = new Map<string, number>();

    for (const item of atendimentos) {
      const key = item.qualidade || 'Indefinido';
      map.set(key, (map.get(key) || 0) + 1);
    }

    return Array.from(map.entries()).map(([name, value]) => ({ name, value }));
  }, [atendimentos]);

  const clientes = useMemo(() => {
    const unique = new Set(
      atendimentos.map((a) => a.numeroCliente).filter(Boolean)
    );
    return unique.size;
  }, [atendimentos]);

  const nav = [
    ['dashboard', BarChart3, 'Dashboard'],
    ['importar', Upload, 'Importar'],
    ['tabela', FileText, 'Tabela Analítica'],
    ['jornada', Users, 'Jornada Cliente'],
    ['qualidade', ShieldCheck, 'Qualidade'],
    ['criticos', AlertTriangle, 'Casos Críticos'],
    ['insights', Brain, 'Insights IA'],
  ] as const;

  const clienteJornada =
    filtered.find((a) => a.numeroCliente)?.numeroCliente ||
    atendimentos.find((a) => a.numeroCliente)?.numeroCliente ||
    '';

  const jornadaCliente = atendimentos
    .filter((a) => a.numeroCliente === clienteJornada)
    .sort((a, b) => (a.dataObj?.getTime() || 0) - (b.dataObj?.getTime() || 0));

  return (
    <div className="app">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-icon">
            <BarChart3 size={26} />
          </div>
          <div>
            <h1>IA Monitor</h1>
            <p>Atendimento, recorrência e satisfação</p>
          </div>
        </div>

        <nav className="menu">
          {nav.map(([id, Icon, label]) => (
            <button
              key={id}
              onClick={() => setPage(id)}
              className={page === id ? 'active' : ''}
            >
              <Icon size={18} />
              {label}
            </button>
          ))}
        </nav>

        <div className="sidebar-footer">
          <strong>Importador real</strong>
          <p>CSV/XLSX com cruzamento por protocolo e cliente.</p>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div>
            <h1>Dashboard de Efetividade Real</h1>
            <p>Cruza protocolo, contato/origem, pontos de bot e satisfação.</p>
          </div>

          <div className="topbar-actions">
            <button className="btn secondary">
              <Filter size={16} />
              Filtros
            </button>
            <button className="btn primary">
              <Download size={16} />
              Exportar
            </button>
          </div>
        </header>

        <div className="content">
          {page === 'dashboard' && (
            <div className="page">
              {atendimentos.length === 0 && (
                <div className="card upload-card">
                  <Upload size={52} />
                  <h3>Importe seus relatórios para começar</h3>
                  <p>
                    Envie juntos os relatórios de Pontos de Acesso, Protocolos
                    Analítico e Pesquisa de Satisfação.
                  </p>
                  <label className="btn primary">
                    Selecionar arquivos
                    <input
                      type="file"
                      multiple
                      accept=".csv,.xlsx,.xls"
                      onChange={handleFiles}
                      style={{ display: 'none' }}
                    />
                  </label>
                </div>
              )}

              {atendimentos.length > 0 && (
                <>
                  <div className="stats-grid">
                    <StatCard
                      icon={FileText}
                      title="Atendimentos"
                      value={total}
                      helper="protocolos únicos importados"
                      tone="neutral"
                      info="Quantidade total de protocolos únicos encontrados nos relatórios importados."
                    />
                    <StatCard
                      icon={Users}
                      title="Clientes únicos"
                      value={clientes}
                      helper="por contato/origem"
                      tone="blue-bg"
                      info="Quantidade de contatos ou origens diferentes identificados nos relatórios."
                    />
                    <StatCard
                      icon={CheckCircle2}
                      title="Solucionados"
                      value={solucionados}
                      helper={`${naoSolucionados} não solucionados`}
                      tone="green-bg"
                      info="Atendimentos marcados como resolvidos com base nos pontos de bot ou status do relatório."
                    />
                    <StatCard
                      icon={RefreshCcw}
                      title="Clientes recorrentes"
                      value={recorrentes}
                      helper="mais de um protocolo"
                      tone="blue-bg"
                      info="Clientes que aparecem em mais de um protocolo, indicando possível retorno ou nova tentativa de atendimento."
                    />
                    <StatCard
                      icon={Upload}
                      title="Transferidos"
                      value={transferidos}
                      helper="transferência para fila"
                      tone="blue-bg"
                      info="Quantidade de protocolos com ponto de bot indicando transferência para atendimento humano ou fila."
                    />
                    <StatCard
                      icon={Clock}
                      title="Retornos até 72h"
                      value={retornos72h}
                      helper="possível recorrência rápida"
                      tone="orange-bg"
                      info="Protocolos em que o mesmo cliente voltou em até 72 horas após o atendimento anterior."
                    />
                    <StatCard
                      icon={AlertTriangle}
                      title="Possíveis falsos positivos"
                      value={falsoPositivo}
                      helper="solucionado, mas retornou"
                      tone="red-bg"
                      info="Casos em que o atendimento foi marcado como solucionado, mas o cliente voltou depois em até 72 horas."
                    />
                    <StatCard
                      icon={Star}
                      title="Nota média"
                      value={notaMedia}
                      helper={`${avaliacoes.length} pesquisas vinculadas`}
                      tone="yellow-bg"
                      info="Média das notas encontradas no relatório de pesquisa de satisfação."
                    />
                    <StatCard
                      icon={TrendingUp}
                      title="CSAT médio"
                      value={csatMedio}
                      helper="base nota 1 a 5"
                      tone="green-bg"
                      info="Percentual médio de satisfação calculado a partir das notas da pesquisa."
                    />
                  </div>

                  <div className="charts-grid">
                    <div className="card chart-card">
                      <SectionTitle
                        icon={BarChart3}
                        title="Atendimentos e recorrência"
                        description="Volume diário comparado com clientes recorrentes"
                      />
                      <div className="chart-box">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={chartDia}>
                            <XAxis dataKey="dia" />
                            <YAxis />
                            <Tooltip />
                            <Bar dataKey="atendimentos" radius={[8, 8, 0, 0]} />
                            <Bar dataKey="recorrentes" radius={[8, 8, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="card chart-card">
                      <SectionTitle
                        icon={Star}
                        title="Evolução da satisfação"
                        description="Nota média diária da pesquisa"
                      />
                      <div className="chart-box">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={chartDia}>
                            <XAxis dataKey="dia" />
                            <YAxis />
                            <Tooltip />
                            <Line
                              type="monotone"
                              dataKey="satisfacao"
                              strokeWidth={3}
                            />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="card chart-card">
                      <SectionTitle
                        icon={AlertTriangle}
                        title="Principais motivos"
                        description="Motivos mais recorrentes no período"
                      />
                      <div className="chart-box">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={motivos} layout="vertical">
                            <XAxis type="number" />
                            <YAxis dataKey="name" type="category" width={170} />
                            <Tooltip />
                            <Bar dataKey="value" radius={[0, 8, 8, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    <div className="card chart-card">
                      <SectionTitle
                        icon={ShieldCheck}
                        title="Qualidade da solução"
                        description="Classificação cruzando solução, retorno e satisfação"
                      />
                      <div className="chart-box">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie
                              data={qualidadeData}
                              dataKey="value"
                              nameKey="name"
                              outerRadius={105}
                              label
                            />
                            <Tooltip />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>
                </>
              )}
            </div>
          )}

          {page === 'importar' && (
            <div className="page">
              <SectionTitle
                icon={Upload}
                title="Importar Relatórios"
                description="Envie relatórios de atendimento/bot e pesquisa de satisfação."
              />

              <div className="card upload-card">
                <Upload size={52} />
                <h3>
                  {isImporting
                    ? 'Importando arquivos...'
                    : 'Selecione os relatórios'}
                </h3>
                <p>
                  Aceita CSV, XLSX e XLS. Pode selecionar os 3 arquivos juntos.
                  O sistema identifica automaticamente o tipo do relatório.
                </p>

                <label className="btn primary">
                  {isImporting ? 'Processando...' : 'Selecionar arquivos'}
                  <input
                    type="file"
                    multiple
                    accept=".csv,.xlsx,.xls"
                    onChange={handleFiles}
                    disabled={isImporting}
                    style={{ display: 'none' }}
                  />
                </label>
              </div>

              {logs.length > 0 && (
                <div className="card table-card">
                  <table>
                    <thead>
                      <tr>
                        <th>Arquivo</th>
                        <th>Tipo identificado</th>
                        <th>Registros lidos</th>
                        <th>Ação</th>
                      </tr>
                    </thead>

                    <tbody>
                      {logs.map((log) => (
                        <tr key={log.id}>
                          <td className="strong">{log.arquivo}</td>
                          <td>{log.tipo}</td>
                          <td>{log.registros}</td>
                          <td>
                            <button
                              className="btn danger tiny"
                              onClick={() => handleDeleteImportedFile(log.id)}
                            >
                              <Trash2 size={15} />
                              Excluir
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {page === 'tabela' && (
            <div className="page">
              <SectionTitle
                icon={FileText}
                title="Tabela Analítica"
                description="Busca por protocolo, cliente, motivo, erro ou ponto de bot."
              />

              <div className="search-box">
                <Search size={18} />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Buscar protocolo, contato, cliente, motivo, erro ou ponto de bot..."
                />
              </div>

              <div className="card table-card">
                <table>
                  <thead>
                    <tr>
                      <th>Protocolo</th>
                      <th>Contato/Cliente</th>
                      <th>Motivo</th>
                      <th>Solução</th>
                      <th>Transferido</th>
                      <th>Recorrência</th>
                      <th>Satisfação</th>
                      <th>Qualidade</th>
                      <th>Próximo protocolo</th>
                      <th>Detalhe</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((a) => (
                      <tr key={a.protocolo}>
                        <td className="strong">{a.protocolo}</td>
                        <td>
                          {a.numeroCliente || '-'}
                          <br />
                          <small>{a.cliente}</small>
                        </td>
                        <td>{a.motivo || '-'}</td>
                        <td>{statusBadge(a)}</td>
                        <td>
                          {a.transferido ? (
                            <Badge tone="blue">Sim</Badge>
                          ) : (
                            <Badge>Não</Badge>
                          )}
                        </td>
                        <td>
                          {a.recorrente ? (
                            <Badge tone="blue">
                              {a.classificacaoRecorrencia}
                            </Badge>
                          ) : (
                            <Badge>Único</Badge>
                          )}
                        </td>
                        <td>{a.nota !== null ? `${a.nota}/5` : '-'}</td>
                        <td>
                          <Badge tone={qualityTone(a.qualidade)}>
                            {a.qualidade}
                          </Badge>
                        </td>
                        <td>{a.proximoProtocolo}</td>
                        <td>
                          <button className="btn tiny">
                            <Eye size={15} />
                            Ver
                          </button>
                        </td>
                      </tr>
                    ))}

                    {filtered.length === 0 && (
                      <tr>
                        <td colSpan={10}>Nenhum registro encontrado.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {page === 'jornada' && (
            <div className="page">
              <SectionTitle
                icon={Users}
                title="Jornada do Cliente"
                description="Visualize todos os protocolos vinculados ao mesmo contato/origem."
              />

              <div className="search-box">
                <Search size={18} />
                <input
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  placeholder="Digite contato, cliente ou protocolo..."
                />
              </div>

              <div className="card">
                <div style={{ padding: 20 }}>
                  <strong>Cliente analisado:</strong>{' '}
                  {clienteJornada || 'nenhum'}
                </div>

                <div className="timeline" style={{ padding: '0 20px 20px' }}>
                  {jornadaCliente.map((a, idx) => (
                    <div key={a.protocolo} className="timeline-item">
                      <div className="timeline-number">{idx + 1}</div>
                      <div>
                        <div className="timeline-header">
                          <strong>{a.protocolo}</strong>
                          {statusBadge(a)}
                          <Badge tone="blue">{a.data || 'Sem data'}</Badge>
                          <Badge tone={qualityTone(a.qualidade)}>
                            {a.qualidade}
                          </Badge>
                        </div>
                        <p>
                          {a.motivo} — {a.erro || 'sem erro identificado'}
                        </p>
                        <small>
                          Próximo contato: {a.proximoProtocolo} | Tempo até
                          retorno: {a.tempoAteProximoContato}
                        </small>
                        {a.comentario && (
                          <p>
                            <small>Comentário: “{a.comentario}”</small>
                          </p>
                        )}
                      </div>
                    </div>
                  ))}

                  {jornadaCliente.length === 0 && (
                    <p style={{ color: '#64748b' }}>
                      Importe dados ou busque um cliente.
                    </p>
                  )}
                </div>
              </div>
            </div>
          )}

          {page === 'qualidade' && (
            <div className="page">
              <SectionTitle
                icon={ShieldCheck}
                title="Qualidade da Solução"
                description="Classificação real considerando solução, retorno e satisfação."
              />

              <div className="quality-grid">
                {qualidadeData.map((item) => (
                  <div key={item.name} className="card quality-card">
                    <p>Status de qualidade</p>
                    <h3>{item.name}</h3>
                    <strong>{item.value}</strong>
                  </div>
                ))}

                {qualidadeData.length === 0 && (
                  <div className="card quality-card">
                    <p>Status de qualidade</p>
                    <h3>Nenhum dado importado</h3>
                    <strong>0</strong>
                  </div>
                )}
              </div>
            </div>
          )}

          {page === 'criticos' && (
            <div className="page">
              <SectionTitle
                icon={AlertTriangle}
                title="Casos Críticos"
                description="Atendimentos com retorno, nota baixa, falso positivo ou erro repetido."
              />

              <div className="critical-list">
                {atendimentos
                  .filter(
                    (a) =>
                      a.retornou72h ||
                      normalizeHeader(a.qualidade).includes('falso') ||
                      normalizeHeader(a.qualidade).includes('questionavel') ||
                      normalizeHeader(a.qualidade).includes('problema')
                  )
                  .map((a) => (
                    <div key={a.protocolo} className="card critical-card">
                      <div>
                        <div className="critical-header">
                          <h3>
                            {a.numeroCliente || a.cliente || 'Sem cliente'}
                          </h3>
                          <Badge tone={qualityTone(a.qualidade)}>
                            {a.qualidade}
                          </Badge>
                          <Badge tone="blue">Protocolo {a.protocolo}</Badge>
                        </div>
                        <p>
                          {a.motivo} — {a.erro || 'sem erro identificado'}
                        </p>
                        <small>
                          Próximo protocolo: {a.proximoProtocolo} | Retorno:{' '}
                          {a.tempoAteProximoContato}
                        </small>
                        {a.comentario && (
                          <p>
                            <small>Comentário: “{a.comentario}”</small>
                          </p>
                        )}
                      </div>
                      <button className="btn secondary">Revisar</button>
                    </div>
                  ))}

                {atendimentos.length === 0 && (
                  <div className="card small-card">
                    <h3>Nenhum arquivo importado</h3>
                    <p>Importe os relatórios para gerar casos críticos.</p>
                  </div>
                )}
              </div>
            </div>
          )}

          {page === 'insights' && (
            <div className="page">
              <SectionTitle
                icon={Brain}
                title="Insights da IA"
                description="Análises operacionais geradas a partir dos dados cruzados."
              />

              <div className="insights-grid">
                <div className="card insight-card">
                  <h3>1. Resumo geral</h3>
                  <p>
                    Foram importados {total} protocolos, com {clientes} clientes
                    únicos,
                    {recorrentes} clientes recorrentes e {retornos72h} retornos
                    em até 72h.
                  </p>
                </div>

                <div className="card insight-card">
                  <h3>2. Falso positivo de solução</h3>
                  <p>
                    Existem {falsoPositivo} protocolos classificados como
                    possível falso positivo. Esses casos foram marcados como
                    solucionados, mas tiveram retorno posterior.
                  </p>
                </div>

                <div className="card insight-card">
                  <h3>3. Satisfação</h3>
                  <p>
                    Foram vinculadas {avaliacoes.length} respostas de
                    satisfação. A nota média atual é {notaMedia} em escala de 1
                    a 5, com CSAT médio de {csatMedio}.
                  </p>
                </div>

                <div className="card insight-card">
                  <h3>4. Motivo mais frequente</h3>
                  <p>
                    O motivo mais frequente no período é “
                    {motivos[0]?.name || 'não identificado'}”, com{' '}
                    {motivos[0]?.value || 0} ocorrências.
                  </p>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
}
