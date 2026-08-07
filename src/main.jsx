import React, { useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import {
  Activity, AlertTriangle, BarChart3, CalendarDays, CheckCircle2,
  ChevronDown, ChevronRight, CircleDot, Clock3, FileSpreadsheet,
  LayoutDashboard, Search, SlidersHorizontal, Sparkles, Users, X
} from 'lucide-react';
import './styles.css';

const today = new Date();
const iso = (offset) => { const d = new Date(today); d.setDate(d.getDate() + offset); return d.toISOString().slice(0, 10); };

const DEMO_TEAMS = [
  { name: 'Produtos Digitais', developers: 7 },
  { name: 'Dados & Analytics', developers: 5 },
  { name: 'Plataformas', developers: 6 },
  { name: 'Experiência do Cliente', developers: 4 },
];
const DEMO_TASKS = [
  ['Novo Portal de Serviços', 'Produtos Digitais', -24, 18, 'Em execução', 62, 'Alta'],
  ['Jornada de autenticação', 'Produtos Digitais', -14, -2, 'Em execução', 78, 'Crítica'],
  ['Catálogo unificado', 'Produtos Digitais', 10, 42, 'Planejada', 0, 'Média'],
  ['Melhoria no checkout', 'Produtos Digitais', -42, -9, 'Concluída', 100, 'Alta'],
  ['Painel de indicadores', 'Dados & Analytics', -8, 27, 'Em execução', 46, 'Alta'],
  ['Qualidade da base cadastral', 'Dados & Analytics', -30, 6, 'Em execução', 82, 'Média'],
  ['Modelo de segmentação', 'Dados & Analytics', 18, 53, 'Planejada', 0, 'Média'],
  ['Migração do data lake', 'Plataformas', -62, -7, 'Em execução', 91, 'Crítica'],
  ['Atualização de APIs legadas', 'Plataformas', -5, 39, 'Em execução', 34, 'Alta'],
  ['Observabilidade de serviços', 'Plataformas', 8, 35, 'Planejada', 0, 'Alta'],
  ['Esteira de CI/CD', 'Plataformas', -53, -12, 'Concluída', 100, 'Média'],
  ['Pesquisa de satisfação', 'Experiência do Cliente', -20, 4, 'Em execução', 72, 'Alta'],
  ['Central de ajuda', 'Experiência do Cliente', 12, 48, 'Planejada', 0, 'Média'],
  ['Revisão da jornada mobile', 'Experiência do Cliente', -39, -5, 'Concluída', 100, 'Alta'],
].map((x, i) => ({ id: i + 1, title: x[0], team: x[1], start: iso(x[2]), end: iso(x[3]), status: x[4], progress: x[5], priority: x[6] }));

const norm = (v) => String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
const first = (row, names) => { const key = Object.keys(row).find(k => names.some(n => norm(k).includes(n))); return key ? row[key] : ''; };
const dateValue = (v) => {
  if (!v) return '';
  if (typeof v === 'number') return XLSX.SSF.parse_date_code(v) ? new Date(Date.UTC(XLSX.SSF.parse_date_code(v).y, XLSX.SSF.parse_date_code(v).m - 1, XLSX.SSF.parse_date_code(v).d)).toISOString().slice(0, 10) : '';
  const d = new Date(v); return Number.isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
};
const normalizeStatus = (value, progress) => {
  const s = norm(value);
  if (s.includes('conclu') || s.includes('complete') || Number(progress) >= 100) return 'Concluída';
  if (s.includes('exec') || s.includes('andamento') || s.includes('progres') || s.includes('iniciad')) return 'Em execução';
  return 'Planejada';
};
const isLate = (t) => t.status !== 'Concluída' && t.end && new Date(`${t.end}T23:59:59`) < today;
const fmt = (v) => v ? new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(`${v}T12:00:00`)).replace('.', '') : 'Sem data';

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const bucketSheet = wb.SheetNames.find(n => norm(n).includes('bucket'));
  let teams = [];
  if (bucketSheet) {
    teams = XLSX.utils.sheet_to_json(wb.Sheets[bucketSheet], { defval: '' }).map(r => ({
      name: String(first(r, ['bucket', 'time', 'equipe', 'nome'])).trim(),
      developers: Number(first(r, ['desenvolvedor', 'quantidade', 'qtd'])) || 0,
    })).filter(t => t.name);
  }
  const candidate = wb.SheetNames.find(n => n !== bucketSheet) || wb.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[candidate], { defval: '', raw: true });
  const tasks = rows.map((r, i) => {
    const progressRaw = first(r, ['progresso', 'percentual', '% conclu', 'conclusao %']);
    const progress = Math.min(100, Number(String(progressRaw).replace('%', '').replace(',', '.')) * (String(progressRaw).includes('%') ? 1 : Number(progressRaw) <= 1 ? 100 : 1) || 0);
    return {
      id: i + 1,
      title: String(first(r, ['titulo', 'tarefa', 'demanda', 'title', 'nome'])).trim(),
      team: String(first(r, ['bucket', 'time', 'equipe'])).trim(),
      start: dateValue(first(r, ['data de inicio', 'inicio', 'start date'])),
      end: dateValue(first(r, ['data de conclusao', 'conclusao', 'termino', 'due date', 'previsao'])),
      status: normalizeStatus(first(r, ['status', 'andamento']), progress), progress,
      priority: String(first(r, ['prioridade', 'priority'])) || 'Não informada',
    };
  }).filter(t => t.title && t.team);
  if (!teams.length) teams = [...new Set(tasks.map(t => t.team))].map(name => ({ name, developers: 0 }));
  return { teams, tasks };
}

const StatusBadge = ({ task }) => {
  const late = isLate(task); const label = late ? 'Atrasada' : task.status;
  return <span className={`status ${late ? 'late' : norm(task.status).replace(' ', '-')}`}><i />{label}</span>;
};

function TeamCard({ team, tasks, open, onToggle }) {
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
          return <div className="timeline-row" key={task.id}><div className="task-name"><strong>{task.title}</strong><small>{task.priority}</small></div><div className="track"><span className="today" style={{left:`${Math.max(0, Math.min(100,(today-min)/span*100))}%`}}/><span className={`bar ${isLate(task)?'late':norm(task.status)}`} style={{left:`${left}%`,width:`${Math.min(width,100-left)}%`}}><i style={{width:`${task.progress}%`}}/></span></div><StatusBadge task={task}/></div>
        })}
      </div>
    </div>}
  </article>;
}

function App() {
  const [teams, setTeams] = useState(DEMO_TEAMS), [tasks, setTasks] = useState(DEMO_TASKS);
  const [query, setQuery] = useState(''), [status, setStatus] = useState('Todos'), [team, setTeam] = useState('Todos');
  const [open, setOpen] = useState(new Set([DEMO_TEAMS[0].name])), [source, setSource] = useState('Dados de demonstração');
  const input = useRef();
  const filtered = useMemo(() => tasks.filter(t => (team === 'Todos' || t.team === team) && (status === 'Todos' || (status === 'Atrasadas' ? isLate(t) : t.status === status && !isLate(t))) && norm(t.title).includes(norm(query))), [tasks, query, status, team]);
  const stats = [
    ['Total de demandas', filtered.length, LayoutDashboard, 'blue'],
    ['Em execução', filtered.filter(t=>t.status==='Em execução'&&!isLate(t)).length, Activity, 'purple'],
    ['Concluídas', filtered.filter(t=>t.status==='Concluída').length, CheckCircle2, 'green'],
    ['Atrasadas', filtered.filter(isLate).length, AlertTriangle, 'red'],
    ['Planejadas', filtered.filter(t=>t.status==='Planejada'&&!isLate(t)).length, Clock3, 'amber'],
  ];
  const load = async e => { const file=e.target.files?.[0]; if(!file)return; try { const data=parseWorkbook(await file.arrayBuffer()); if(!data.tasks.length) throw new Error('Nenhuma demanda encontrada'); setTeams(data.teams);setTasks(data.tasks);setSource(file.name);setOpen(new Set(data.teams.slice(0,1).map(t=>t.name))); } catch(err) { alert(`Não foi possível ler a planilha: ${err.message}`); } e.target.value=''; };
  const clear = () => { setQuery(''); setStatus('Todos'); setTeam('Todos'); };
  const visibleTeams = teams.filter(t => team === 'Todos' || t.name === team).filter(t => filtered.some(d => d.team === t.name));
  return <div className="app">
    <header><div className="brand"><div className="brand-mark"><BarChart3/></div><div><strong>CDES</strong><span>Painel de Demandas</span></div></div><div className="header-right"><span className="updated"><CircleDot size={14}/> Atualizado agora</span><button className="import" onClick={()=>input.current.click()}><FileSpreadsheet size={17}/> Importar Excel</button><input ref={input} type="file" accept=".xlsx,.xls" hidden onChange={load}/><div className="avatar">GE</div></div></header>
    <main>
      <section className="hero"><div><div className="eyebrow"><Sparkles size={14}/> VISÃO EXECUTIVA</div><h1>Painel de Demandas</h1><p>Acompanhe o portfólio, identifique riscos e veja a evolução de cada time.</p></div><div className="source"><span>Fonte de dados</span><strong><FileSpreadsheet size={16}/>{source}</strong></div></section>
      <section className="filters">
        <label className="search"><Search size={19}/><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Pesquisar por demanda..." />{query&&<button onClick={()=>setQuery('')}><X size={16}/></button>}</label>
        <label><Users size={17}/><select value={team} onChange={e=>setTeam(e.target.value)}><option>Todos</option>{teams.map(t=><option key={t.name}>{t.name}</option>)}</select></label>
        <label><SlidersHorizontal size={17}/><select value={status} onChange={e=>setStatus(e.target.value)}><option>Todos</option><option>Planejada</option><option>Em execução</option><option>Concluída</option><option>Atrasadas</option></select></label>
        {(query||team!=='Todos'||status!=='Todos')&&<button className="clear" onClick={clear}><X size={15}/> Limpar filtros</button>}
      </section>
      <section className="stats">{stats.map(([label,value,Icon,color])=><div className="stat" key={label}><div className={`stat-icon ${color}`}><Icon/></div><div><span>{label}</span><strong>{value}</strong></div>{label==='Atrasadas'&&value>0&&<em>Requer atenção</em>}</div>)}</section>
      <section className="teams-section"><div className="section-title"><div><h2>Visão por time</h2><p>{visibleTeams.length} times com demandas no período</p></div><button onClick={()=>setOpen(open.size===visibleTeams.length?new Set():new Set(visibleTeams.map(t=>t.name)))}>{open.size===visibleTeams.length?'Recolher todos':'Expandir todos'}</button></div>
        <div className="team-list">{visibleTeams.map(t=><TeamCard key={t.name} team={t} tasks={filtered.filter(d=>d.team===t.name)} open={open.has(t.name)} onToggle={()=>setOpen(s=>{const n=new Set(s);n.has(t.name)?n.delete(t.name):n.add(t.name);return n;})}/>)}{!visibleTeams.length&&<div className="empty"><Search/><h3>Nenhuma demanda encontrada</h3><p>Tente ajustar os filtros ou a pesquisa.</p><button onClick={clear}>Limpar filtros</button></div>}</div>
      </section>
    </main><footer>CDES · Gestão de Portfólio <span>•</span> Dados consolidados do Microsoft Planner</footer>
  </div>;
}

createRoot(document.getElementById('root')).render(<App/>);
