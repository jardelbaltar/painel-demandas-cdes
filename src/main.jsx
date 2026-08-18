import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import {
  Activity, AlertTriangle, BarChart3, CalendarDays, CheckCircle2,
  Check, ChevronDown, ChevronRight, CircleDot, Clock3, FileSpreadsheet,
  LayoutDashboard, LogIn, RefreshCw, Search, SlidersHorizontal, Sparkles, Users, X
} from 'lucide-react';
import * as XLSX from 'xlsx';
import './styles.css';
import defaultData from './default-data.json';
import { isPlannerConfigured, loadPlanner, parseBucketName } from './planner.js';

const today = new Date();

const norm = (v) => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
const first = (row, names) => {
  const keys = Object.keys(row);
  const exact = names.map(norm).map(name => keys.find(key => norm(key) === name)).find(Boolean);
  const key = exact || keys.find(k => names.some(n => norm(k).includes(norm(n))));
  return key ? row[key] : '';
};
const dateValue = (v) => {
  if (!v) return '';
  if (typeof v === 'number') return XLSX.SSF.parse_date_code(v) ? new Date(Date.UTC(XLSX.SSF.parse_date_code(v).y, XLSX.SSF.parse_date_code(v).m - 1, XLSX.SSF.parse_date_code(v).d)).toISOString().slice(0, 10) : '';
  const d = new Date(v); return Number.isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
};
const normalizeStatus = (value, progress) => {
  const s = norm(value);
  if (s.includes('conclu') || s.includes('complete') || Number(progress) >= 100) return 'Concluída';
  if (s.includes('nao iniciad')) return 'Planejada';
  if (s.includes('exec') || s.includes('andamento') || s.includes('progres') || s.includes('iniciad')) return 'Em execução';
  return 'Planejada';
};
const highlightedLabel = (value) => String(value ?? '').split(/[;,|]/)
  .map(label => label.trim())
  .find(label => ['suspensa', 'em homologacao'].includes(norm(label))) || '';
const isLate = (t) => t.status !== 'Concluída' && t.end && new Date(`${t.end}T23:59:59`) < today;
const fmt = (v) => v ? new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(`${v}T12:00:00`)).replace('.', '') : 'Sem data';
const numericDate = value => {
  if (value instanceof Date) return value;
  const match = String(value ?? '').trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);
  return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])) : null;
};
const detailValue = (label, value) => {
  const date = numericDate(value);
  if (date && (value instanceof Date || ['data', 'criado em', 'concluido em'].some(name => norm(label).includes(name)))) {
    return new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' }).format(date);
  }
  if (typeof value === 'boolean') return value ? 'Sim' : 'Não';
  return String(value ?? '').trim();
};
const checklistItems = (value, completedValue) => {
  const completedCount = Number(String(completedValue ?? '').match(/^\s*(\d+)/)?.[1]) || 0;
  return String(value ?? '').split(/[;\n]+/).map(title => title.trim()).filter(Boolean)
    .map((title, index) => ({ title, completed: index < completedCount }));
};

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const bucketSheet = wb.SheetNames.find(n => norm(n).includes('bucket'));
  let teams = [];
  if (bucketSheet) {
    teams = XLSX.utils.sheet_to_json(wb.Sheets[bucketSheet], { defval: '' }).map(r => {
      const parsed = parseBucketName(first(r, ['nome do bucket', 'nome', 'bucket', 'time', 'equipe']));
      return { ...parsed, developers: Number(first(r, ['desenvolvedor', 'quantidade', 'qtd'])) || parsed.developers };
    }).filter(t => t.name);
  }
  const candidate = wb.SheetNames.find(n => norm(n).includes('dados consolidados'))
    || wb.SheetNames.find(n => norm(n) === 'tarefas')
    || wb.SheetNames.find(n => n !== bucketSheet);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[candidate], { defval: '', raw: true });
  const tasks = rows.map((r, i) => {
    const progressRaw = first(r, ['progresso', 'percentual', '% conclu', 'conclusao %']);
    const progress = Math.min(100, Number(String(progressRaw).replace('%', '').replace(',', '.')) * (String(progressRaw).includes('%') ? 1 : Number(progressRaw) <= 1 ? 100 : 1) || 0);
    return {
      id: i + 1,
      title: String(first(r, ['nome da tarefa', 'nome da demanda', 'titulo', 'tarefa', 'demanda', 'title', 'nome'])).trim(),
      team: parseBucketName(first(r, ['categoria', 'nome do bucket', 'bucket', 'time', 'equipe'])).name,
      start: dateValue(first(r, ['data de inicio', 'inicio', 'start date'])),
      end: dateValue(first(r, ['data de conclusao', 'conclusao', 'termino', 'due date', 'previsao'])),
      status: normalizeStatus(first(r, ['status', 'andamento']), progress), progress,
      priority: String(first(r, ['prioridade', 'priority'])) || 'Não informada',
      highlightedLabel: highlightedLabel(first(r, ['rotulos', 'rotulo', 'labels', 'label'])),
      checklist: checklistItems(first(r, ['itens da lista de verificacao']), first(r, ['itens concluidos da lista de verificacao'])),
      details: Object.entries(r).map(([label, value]) => ({ label, value: detailValue(label, value) })).filter(detail => detail.value),
    };
  }).filter(t => t.title && t.team);
  if (!teams.length) teams = [...new Set(tasks.map(t => t.team))].map(name => ({ name, developers: 0 }));
  return { teams, tasks };
}

const StatusBadge = ({ task }) => {
  const late = isLate(task); const label = late ? 'Atrasada' : task.status;
  return <span className={`status ${late ? 'late' : norm(task.status).replace(' ', '-')}`}><i />{label}</span>;
};

function DemandModal({ task, onClose }) {
  const dialog = useRef(null);
  useEffect(() => {
    const previousOverflow = document.body.style.overflow;
    const onKeyDown = event => event.key === 'Escape' && onClose();
    document.body.style.overflow = 'hidden';
    document.addEventListener('keydown', onKeyDown);
    dialog.current?.focus();
    return () => { document.body.style.overflow = previousOverflow; document.removeEventListener('keydown', onKeyDown); };
  }, [onClose]);
  const hiddenDetails = new Set(['identificacao da tarefa', 'categoria', 'itens concluidos da lista de verificacao', 'atrasados', 'itens da lista de verificacao']);
  const details = (task.details?.length ? task.details : [
    { label: 'Time responsável', value: task.team }, { label: 'Status', value: isLate(task) ? 'Atrasada' : task.status },
    { label: 'Prioridade', value: task.priority }, { label: 'Progresso', value: `${task.progress}%` },
    { label: 'Data de início', value: task.start ? detailValue('Data de início', task.start) : 'Sem data' },
    { label: 'Data de conclusão', value: task.end ? detailValue('Data de conclusão', task.end) : 'Sem data' },
    ...(task.highlightedLabel ? [{ label: 'Rótulo', value: task.highlightedLabel }] : []),
  ]).filter(detail => !hiddenDetails.has(norm(detail.label)));
  return <div className="modal-backdrop" onMouseDown={onClose} role="presentation">
    <section className="demand-modal" role="dialog" aria-modal="true" aria-labelledby="demand-modal-title" tabIndex={-1} ref={dialog} onMouseDown={event => event.stopPropagation()}>
      <div className="modal-heading"><span>DETALHES DA DEMANDA</span><h2 id="demand-modal-title">{task.title}</h2><p>{task.team}</p></div>
      <div className="modal-summary"><StatusBadge task={task}/><span className="modal-progress"><i style={{ width: `${task.progress}%` }}/></span><strong>{task.progress}%</strong></div>
      <div className="detail-grid">{details.map((detail, index) => <div className="detail-item" key={`${detail.label}-${index}`}><span>{detail.label}</span><strong>{detail.value}</strong></div>)}</div>
      {task.checklist?.length > 0 && <section className="modal-checklist"><h3>Lista de verificação</h3><ul>{task.checklist.map((item, index) => <li className={item.completed ? 'completed' : ''} key={`${item.title}-${index}`}><span className="checklist-icon">{item.completed && <Check size={14} strokeWidth={3}/>}</span><span>{item.title}</span></li>)}</ul></section>}
      <p className="modal-hint">Clique fora desta janela para fechar</p>
    </section>
  </div>;
}

function TeamCard({ team, tasks, open, onToggle, onSelectTask }) {
  const counts = { planned: tasks.filter(t => t.status === 'Planejada' && !isLate(t)).length, active: tasks.filter(t => t.status === 'Em execução' && !isLate(t)).length, late: tasks.filter(isLate).length, done: tasks.filter(t => t.status === 'Concluída').length };
  const dated = tasks.filter(t => t.start && t.end);
  const min = dated.length ? Math.min(...dated.map(t => new Date(t.start))) : +today;
  const max = dated.length ? Math.max(...dated.map(t => new Date(t.end)), +today) : +today + 86400000;
  const span = Math.max(max - min, 86400000);
  const monthLabels = [0, .25, .5, .75, 1].map(p => new Date(min + span * p));
  return <article className={`team-card ${open ? 'expanded' : ''}`}>
    <button className="team-head" onClick={onToggle} aria-expanded={open}>
      <div className="team-identity"><div className="team-icon"><Users size={20} /></div><div><h3>{team.name}</h3><p>{team.developers || '—'} desenvolvedores <span>•</span> {tasks.length} demandas</p></div></div>
      <div className="team-chips">
        <span className="chip planned"><b>{counts.planned}</b> Planejadas</span><span className="chip active"><b>{counts.active}</b> Em execução</span>
        <span className="chip late"><b>{counts.late}</b> Atrasadas</span><span className="chip done"><b>{counts.done}</b> Concluídas</span>
      </div><span className="expand">{open ? <ChevronDown /> : <ChevronRight />}</span>
    </button>
    {open && <div className="roadmap-wrap">
      <div className="roadmap-title"><div><h4>Roadmap de demandas</h4><p>Cronograma previsto e progresso das entregas</p></div><span><CalendarDays size={16}/> {fmt(new Date(min).toISOString().slice(0,10))} — {fmt(new Date(max).toISOString().slice(0,10))}</span></div>
      <div className="roadmap">
        <div className="timeline-head"><span>Demanda</span><div>{monthLabels.map((d,i)=><small key={i}>{new Intl.DateTimeFormat('pt-BR',{month:'short'}).format(d).replace('.','')}</small>)}</div><span>Status</span></div>
        {tasks.map(task => {
          const left = task.start ? Math.max(0, (new Date(task.start)-min)/span*100) : 0;
          const width = task.end && task.start ? Math.max(3, (new Date(task.end)-new Date(task.start))/span*100) : 3;
          return <div className="timeline-row" key={task.id}><div className="task-name"><button className="task-title" onClick={() => onSelectTask(task)}>{task.title}</button>{task.highlightedLabel && <span className={`task-label ${norm(task.highlightedLabel)}`}>{task.highlightedLabel}</span>}<small>{task.priority}</small></div><div className="track"><span className="today" style={{left:`${Math.max(0, Math.min(100,(today-min)/span*100))}%`}}/><span className={`bar ${isLate(task)?'late':norm(task.status)}`} style={{left:`${left}%`,width:`${Math.min(width,100-left)}%`}}><i style={{width:`${task.progress}%`}}/></span></div><StatusBadge task={task}/></div>
        })}
      </div>
    </div>}
  </article>;
}

function App() {
  const [teams, setTeams] = useState([]), [tasks, setTasks] = useState([]);
  const [query, setQuery] = useState(''), [status, setStatus] = useState('Todos'), [team, setTeam] = useState('Todos');
  const [open, setOpen] = useState(new Set());
  const [dataSource, setDataSource] = useState({ type: 'loading', label: 'Carregando dados…', detail: '' });
  const [syncing, setSyncing] = useState(false);
  const [selectedTask, setSelectedTask] = useState(null);
  const applyWorkbook = (buffer, fileName) => {
    const data = parseWorkbook(buffer);
    if (!data.tasks.length) throw new Error('Nenhuma demanda encontrada');
    setTeams(data.teams); setTasks(data.tasks);
    setDataSource({ type: 'spreadsheet', label: 'Planilha', detail: fileName });
    setOpen(new Set(data.teams.slice(0, 1).map(t => t.name)));
  };
  const loadSpreadsheetFallback = async plannerError => {
    try {
      const response = await fetch('./produtos-e-times.xlsx');
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      applyWorkbook(await response.arrayBuffer(), 'produtos-e-times.xlsx');
      setDataSource({ type: 'fallback', label: 'Planilha (modo contingência)', detail: plannerError });
    } catch (error) {
      console.error('Falha ao carregar a planilha publicada:', error);
      setTeams(defaultData.teams); setTasks(defaultData.tasks);
      setOpen(new Set(defaultData.teams.slice(0, 1).map(t => t.name)));
      setDataSource({ type: 'fallback', label: 'Dados incluídos no painel', detail: plannerError });
    }
  };
  const syncPlanner = async interactive => {
    setSyncing(true);
    try {
      const data = await loadPlanner({ interactive });
      setTeams(data.teams); setTasks(data.tasks);
      setOpen(new Set(data.teams.slice(0, 1).map(t => t.name)));
      setDataSource({ type: 'planner', label: 'Sincronizado com o Planner', detail: `${data.account?.name || data.account?.preferred_username || 'Usuário Microsoft'} · ${data.syncedAt.toLocaleString('pt-BR')}` });
    } catch (error) {
      console.warn('Não foi possível sincronizar com o Planner:', error);
      const reason = !isPlannerConfigured() ? 'Integração Microsoft não configurada' : 'Autenticação ou acesso ao plano indisponível';
      await loadSpreadsheetFallback(reason);
    } finally { setSyncing(false); }
  };
  useEffect(() => { syncPlanner(false); }, []);
  const filtered = useMemo(() => tasks.filter(t => (team === 'Todos' || t.team === team) && (status === 'Todos' || (status === 'Atrasadas' ? isLate(t) : t.status === status && !isLate(t))) && norm(t.title).includes(norm(query))), [tasks, query, status, team]);
  const stats = [
    ['Total de demandas', filtered.length, LayoutDashboard, 'blue'],
    ['Em execução', filtered.filter(t=>t.status==='Em execução'&&!isLate(t)).length, Activity, 'purple'],
    ['Concluídas', filtered.filter(t=>t.status==='Concluída').length, CheckCircle2, 'green'],
    ['Atrasadas', filtered.filter(isLate).length, AlertTriangle, 'red'],
    ['Planejadas', filtered.filter(t=>t.status==='Planejada'&&!isLate(t)).length, Clock3, 'amber'],
  ];
  const clear = () => { setQuery(''); setStatus('Todos'); setTeam('Todos'); };
  const visibleTeams = teams.filter(t => team === 'Todos' || t.name === team).filter(t => filtered.some(d => d.team === t.name));
  return <div className="app">
    <header><div className="brand"><div className="brand-mark"><BarChart3/></div><div><strong>CDES</strong><span>Painel de Demandas</span></div></div><div className="header-right"><button className="sync-button" onClick={()=>syncPlanner(true)} disabled={syncing || !isPlannerConfigured()}>{syncing ? <RefreshCw className="spinning" size={17}/> : <LogIn size={17}/>} {syncing ? 'Sincronizando…' : 'Conectar ao Planner'}</button></div></header>
    <main>
      <section className="hero"><div><div className="eyebrow"><Sparkles size={14}/> VISÃO EXECUTIVA</div><h1>Painel de Demandas</h1></div><div className={`source ${dataSource.type}`}><span>Fonte de dados</span><strong>{dataSource.type === 'planner' ? <CircleDot size={16}/> : <FileSpreadsheet size={16}/>} {dataSource.label}</strong>{dataSource.detail && <small>{dataSource.detail}</small>}</div></section>
      <section className="filters">
        <label className="search"><Search size={19}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Pesquisar por demanda..." />{query&&<button onClick={()=>setQuery('')}><X size={16}/></button>}</label>
        <label><Users size={17}/><select value={team} onChange={e=>setTeam(e.target.value)}><option>Todos</option>{teams.map(t=><option key={t.name}>{t.name}</option>)}</select></label>
        <label><SlidersHorizontal size={17}/><select value={status} onChange={e=>setStatus(e.target.value)}><option>Todos</option><option>Planejada</option><option>Em execução</option><option>Concluída</option><option>Atrasadas</option></select></label>
        {(query||team!=='Todos'||status!=='Todos')&&<button className="clear" onClick={clear}><X size={15}/> Limpar filtros</button>}
      </section>
      <section className="stats">{stats.map(([label,value,Icon,color])=><div className="stat" key={label}><div className={`stat-icon ${color}`}><Icon/></div><div><span>{label}</span><strong>{value}</strong></div>{label==='Atrasadas'&&value>0&&<em>Requer atenção</em>}</div>)}</section>
      <section className="teams-section"><div className="section-title"><div><h2>Visão por time</h2><p>{visibleTeams.length} times com demandas no período</p></div><button onClick={()=>setOpen(open.size===visibleTeams.length?new Set():new Set(visibleTeams.map(t=>t.name)))}>{open.size===visibleTeams.length?'Recolher todos':'Expandir todos'}</button></div>
        <div className="team-list">{visibleTeams.map(t=><TeamCard key={t.name} team={t} tasks={filtered.filter(d=>d.team===t.name)} open={open.has(t.name)} onSelectTask={setSelectedTask} onToggle={()=>setOpen(s=>{const n=new Set(s);n.has(t.name)?n.delete(t.name):n.add(t.name);return n;})}/>)}{!visibleTeams.length&&<div className="empty"><Search/><h3>Nenhuma demanda encontrada</h3><p>Tente ajustar os filtros ou a pesquisa.</p><button onClick={clear}>Limpar filtros</button></div>}</div>
      </section>
    </main><footer>CDES · Gestão de Portfólio <span>•</span> Dados consolidados do Microsoft Planner</footer>{selectedTask && <DemandModal task={selectedTask} onClose={() => setSelectedTask(null)}/>}
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
