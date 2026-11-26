/* Visualizações do projeto Escolha sua banca. JavaScript puro, sem bibliotecas. */
'use strict';
const $ = id => document.getElementById(id);
const NS = 'http://www.w3.org/2000/svg';
const fmt = n => Number(n||0).toLocaleString('pt-BR');
const esc = s => String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ROLES = ['Orientação','Coorientação','Titular','Suplente'];
const PAL = ['#3a756e','#d3654a','#d4b66a','#6f8b81','#9e765d','#b4aa91','#83978b'];
const MS_COLORS = {'MS1':'#c6bbb0','MS2':'#9c9b8c','MS3':'#3b756e','MS5':'#d3654a','MS6':'#d4b66a'};
const STATUS_COLORS = {'Ativos':'#3b756e','Aposentados':'#bf8a59','Ex-professores':'#8f8b96','Colaboradores':'#d3654a','In Memoriam':'#3d4850'};
const RACE_COLORS = {'Branca':'#e3d2b1','Branco':'#e3d2b1','Preta':'#313a36','Preto':'#313a36','Parda':'#af6f53','Pardo':'#af6f53','Pardo?':'#ad9d72','Não Declarado':'#909f96','Amarela':'#d9b044','Indígena':'#4d8d7f'};
let DATA = null;
let DOC_ANO = '2026', TREND = 'total';
let FILTER = {ano:'todos',papel:'todos',status:'todos',busca:''};
let SELECTED = null, PREMIO_PERSON = null, PREMIO_ANO = 'todos';
let TROPHY = 0, TROPHY_REQUEST = null;

function showError(message){$('loading').innerHTML=`<span>!</span><p style="max-width:500px;padding:15px;text-align:center">${esc(message)}</p>`;}
function closeLoading(){const el=$('loading');el.classList.add('hide');setTimeout(()=>el.remove(),500);}
function notify(message){const t=$('toast');t.textContent=message;t.classList.add('active');setTimeout(()=>t.classList.remove('active'),4200);}
function initials(name){return String(name).trim().split(/\s+/).slice(0,2).map(x=>x[0]).join('').toUpperCase();}
function makeSVG(tag,attrs={},text){const el=document.createElementNS(NS,tag);Object.entries(attrs).forEach(([k,v])=>el.setAttribute(k,v));if(text!==undefined)el.textContent=text;return el;}
function ordered(items){return Object.entries(items).filter(([_,v])=>Number(v)>0).sort((a,b)=>Number(b[1])-Number(a[1])||a[0].localeCompare(b[0]));}
function bars(id,items,options={}){
  const host=$(id), rows=ordered(items||{}),max=Math.max(1,...rows.map(r=>r[1]));
  host.innerHTML=rows.length?rows.map(([name,count],i)=>`<div class="bar-item"><span class="text">${esc(name)}</span><span class="bar-track"><i style="width:${Math.max(0,100*count/max).toFixed(1)}%;background:${options.colors?.[name]||PAL[i%PAL.length]}"></i></span><strong>${fmt(count)}</strong></div>`).join(''):'<p class="small-note">Sem registros disponíveis para este recorte.</p>';
}
function raceRows(id,items){
  const host=$(id);if(!items){host.innerHTML='<p class="small-note">O Anuário não apresenta raça/cor para este ano.</p>';return;}
  const all=ordered(items);host.innerHTML=all.length?all.map(([name,n])=>`<div class="race-row"><span class="left"><i class="race-dot" style="background:${RACE_COLORS[name]||'#85978f'}"></i>${esc(name)}</span><b>${fmt(n)}</b></div>`).join(''):'<p class="small-note">Sem valores informados.</p>';
}
function heroDots(){
  const host=$('hero-pontos');if(!host)return;
  for(let ring=0;ring<3;ring++)for(let k=0;k<16+ring*8;k++){
    const n=16+ring*8, theta=k*2*Math.PI/n-Math.PI/2,rr=20+ring*16;
    const dot=document.createElement('span');dot.className='orbital-dot'+((k+ring)%7===0?' alt':'');
    dot.style.left=`calc(${50+Math.cos(theta)*rr}% - 5px)`;
    dot.style.top=`calc(${50+Math.sin(theta)*rr}% - 5px)`;
    host.appendChild(dot);
  }
}

// 02. DOCENTES — 2015–2025 vêm do Anuário; 2026 vem de outra fonte.
function yearData(y=DOC_ANO){return DATA.docentes.anos[String(y)];}
function renderDocentes(){
  const d=yearData(),prev=yearData(Number(DOC_ANO)-1);
  $('doc-ano-num').textContent=DOC_ANO;
  $('doc-total').textContent=fmt(d.total);
  const delta=prev?d.total-prev.total:0;
  $('doc-variacao').textContent=prev?`${delta>0?'+':''}${delta} em relação a ${prev.ano}${DOC_ANO==='2026'?' (fontes distintas)':''}`:'Primeiro ano do recorte disponível';
  $('doc-fonte-label').textContent=d.fonte_tipo;
  $('doc-fonte-texto').textContent=d.fonte_titulo;
  $('doc-fonte-link').href=d.fonte_url;
  const regimes=['Integral','Completo','Parcial'];
  $('doc-regimes').innerHTML=regimes.map(r=>`<div class="regime-row"><span>${r}</span><span class="track"><i style="width:${d.total?100*d.regimes[r]/d.total:0}%"></i></span><b>${fmt(d.regimes[r])}</b></div>`).join('');
  bars('doc-ms',d.ms,{colors:MS_COLORS});
  const individual=!!d.genero;
  $('doc-genero-titulo').textContent=individual?'Gênero no levantamento':'Sexo tabulado';
  bars('doc-genero',individual?d.genero:{'Homens (M)':d.sexo.M,'Mulheres (F)':d.sexo.F},
       {colors:{'Homens (M)':'#3b756e','Mulheres (F)':'#d3654a','Homem cis':'#3b756e','Mulher cis':'#d3654a'}});
  $('doc-genero-nota').textContent=individual?'2026: classificação nominal produzida pelos autores do projeto.':`Anuário: sexo tabulado (M/F). Registros preenchidos: ${d.sexo.registrados} de ${d.total}.`;
  raceRows('doc-raca',d.raca);
  $('doc-raca-nota').textContent=individual?'A categoria Pardo? permanece separada, pois a classificação está pendente de verificação.':'As categorias seguem o Anuário. ';
  $('doc-nota-texto').textContent=d.observacoes.length?d.observacoes.join(' '):'Os valores deste ano são os que constam das abas do Anuário. Sexo é uma categoria agregada da fonte, não identidade de gênero individual.';
  renderTimeline();
}
// A legenda usa as MESMAS cores que os segmentos do gráfico, inclusive em 2026.
function renderTimelineLegend(){
  const legend = $('doc-timeline-legend');
  const items = {
    total: [['Docentes ativos','#3b756e']],
    ms: Object.entries(MS_COLORS),
    genero: [
      ['M no Anuário / homem cis no levantamento de 2026','#3b756e'],
      ['F no Anuário / mulher cis no levantamento de 2026','#d3654a']
    ],
    raca: [
      ['Branca / Branco',RACE_COLORS.Branca],
      ['Preta / Preto',RACE_COLORS.Preta],
      ['Parda / Pardo',RACE_COLORS.Parda],
      ['Pardo? (2026, a conferir)',RACE_COLORS['Pardo?']],
      ['Amarela',RACE_COLORS.Amarela],
      ['Indígena',RACE_COLORS.Indígena],
      ['Não Declarado',RACE_COLORS['Não Declarado']]
    ]
  };
  legend.innerHTML=items[TREND].map(([label,color])=>
    `<span class="timeline-key" role="listitem"><i class="timeline-swatch" style="background:${color}"></i>${esc(label)}</span>`
  ).join('');
}
function renderTimeline(){
  const years=DATA.docentes.anos;
  const y=Object.keys(years).map(k=>years[k]);
  const configs={
    total:{caption:'Totais de docentes ativos. O ano de 2026 deriva do levantamento do projeto e não integra o Anuário.',get:d=>[['Docentes ativos',d.total,'#3b756e']],total:d=>d.total},
    ms:{caption:'Classes MS: subclasses 3.1/3.2 e 5.1/5.2/5.3 somadas nos níveis MS3 e MS5.',get:d=>Object.entries(d.ms).map(([k,v])=>[k,v,MS_COLORS[k]||'#93978e']),total:d=>d.total},
    genero:{caption:'De 2015 a 2025: sexo tabulado M/F do Anuário. Em 2026: categorias de gênero do levantamento. As medidas não são diretamente equivalentes.',get:d=>d.genero?Object.entries(d.genero).map(([k,v])=>[k,v,k.toLowerCase().includes('mulher')?'#d3654a':'#3b756e']):[['M',d.sexo.M,'#3b756e'],['F',d.sexo.F,'#d3654a']],total:d=>d.genero?Object.values(d.genero).reduce((a,b)=>a+b,0):d.sexo.registrados},
    raca:{caption:'Raça/cor não está disponível no Anuário de 2015 a 2022. O retrato de 2026 contém classificações pendentes e não constitui uma continuação oficial da série.',get:d=>d.raca?Object.entries(d.raca).map(([k,v])=>[k,v,RACE_COLORS[k]||'#85978f']):[],total:d=>d.raca?Object.values(d.raca).reduce((a,b)=>a+b,0):0}
  };
  const cfg=configs[TREND],max=Math.max(1,...y.map(d=>cfg.total(d)));
  $('doc-timeline').innerHTML=y.map(d=>{
    const parts=cfg.get(d).filter(([_,n])=>n>0),count=cfg.total(d),height=Math.max(1,count/max*150);
    const stacks=parts.map(([n,v,c])=>`<span class="stack" title="${esc(n)}: ${v}" style="height:${count?v/count*100:0}%;background:${c}"></span>`).join('');
    return `<button class="trend-year ${DOC_ANO===String(d.ano)?'current':''} ${d.ano===2026?'provisional':''}" type="button" data-chart-year="${d.ano}" title="${d.ano}: ${count||'sem dados para a categoria'}. Clique para selecionar."><strong>${count||'s/d'}</strong><span class="col-outer"><span class="col-inner" style="height:${count?height:2}px">${stacks}</span></span><span class="year-text">${d.ano}${d.ano===2026?'*':''}</span></button>`;
  }).join('');
  renderTimelineLegend();
  $('doc-timeline-note').textContent=cfg.caption+' O asterisco (*) distingue 2026, elaborado a partir de outra fonte.';
  $('doc-timeline').querySelectorAll('[data-chart-year]').forEach(el=>el.addEventListener('click',()=>{
    DOC_ANO=el.dataset.chartYear;$('ano-docentes').value=DOC_ANO;renderDocentes();
  }));
}
function initDocentes(){
  const selector=$('ano-docentes');for(const year of DATA.docentes.anosDisponiveis){const o=new Option(year,year);selector.add(o);}selector.value=DOC_ANO;
  selector.addEventListener('change',()=>{DOC_ANO=selector.value;renderDocentes();});
  $('doc-tendencia').addEventListener('change',e=>{TREND=e.target.value;renderTimeline();});
  $('doc-status').innerHTML=ordered(DATA.docentes.statusCadastro2026).map(([name,n])=>`<span class="status-chip"><span>${esc(name)}</span><b>${fmt(n)}</b></span>`).join('');
  renderDocentes();
}

// 03. BANCAS — uma bolinha é uma pessoa, independentemente do número de funções.
function getProfessor(id){return DATA.professores.find(p=>p.id===id);}
function participationCounts(p){
  if(FILTER.ano==='todos')return p.funcoes;
  const subset=p.participacoes.filter(r=>String(r.ano)===FILTER.ano);
  return Object.fromEntries(ROLES.map(r=>[r,subset.filter(x=>x.funcao===r).length]));
}
function countRole(p){const c=participationCounts(p);return FILTER.papel==='todos'?ROLES.reduce((n,k)=>n+c[k],0):c[FILTER.papel];}
function mappedPeople(){return DATA.professores.filter(p=>(FILTER.status==='todos'||FILTER.status===p.status)&&p.nome.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').includes(FILTER.busca));}
function renderBancas(){
  const subset=mappedPeople(),n=subset.reduce((a,p)=>a+countRole(p),0),peopleActive=subset.filter(p=>countRole(p)>0).length;
  $('banca-funcoes').textContent=fmt(n);$('banca-pessoas').textContent=fmt(peopleActive);
  $('banca-defesas').textContent=fmt(FILTER.ano==='todos'?DATA.metricas.defesas:DATA.defesas[FILTER.ano].total);
  if(!subset.find(p=>p.id===SELECTED))SELECTED=(subset.find(p=>countRole(p)>0)||subset[0])?.id||null;
  drawPeopleMap(subset);renderSelected();
  // A demografia vem agregada do servidor e não é associada a nomes no navegador.
  const agg=DATA.distribuicoes.bancas[FILTER.ano];
  bars('banca-dist-genero',agg.genero);
  bars('banca-dist-raca',agg.raca);
  bars('banca-dist-status',agg.status,{colors:STATUS_COLORS});
}
// Layout inspirado em mapas parlamentares de assentos: anéis semicirculares,
// lugares equidistantes e POSIÇÕES FIXAS para permitir comparar filtros.
// Não representa hierarquia, proximidade ou desempenho acadêmico.
function seatsHemicircle(total){
  const cx=460, cy=452;
  const radii=[112,152,192,232,272,312,352];
  const base=[6,9,12,14,16,18,20];
  const baseTotal=base.reduce((a,b)=>a+b,0);
  const slots=base.map(x=>Math.floor(total*x/baseTotal));
  let rest=total-slots.reduce((a,b)=>a+b,0);
  // Arredondamento distribuído de modo determinístico; 95 docentes => 95 assentos.
  const remainder=base.map((v,i)=>({i,f:total*v/baseTotal-slots[i]})).sort((a,b)=>b.f-a.f||a.i-b.i);
  for(let k=0;k<rest;k++)slots[remainder[k%remainder.length].i]++;
  const seats=[];
  slots.forEach((count,ring)=>{
    for(let i=0;i<count;i++){
      const theta=Math.PI-Math.PI*(i+.5)/count;
      seats.push({x:cx+Math.cos(theta)*radii[ring],y:cy-Math.sin(theta)*radii[ring],ring});
    }
  });
  return seats.sort((a,b)=>a.x-b.x||a.y-b.y);
}
function drawPeopleMap(matched){
  const svg=$('mapa-docentes');svg.replaceChildren();
  // Todos os professores aparecem, mesmo quando o filtro só destaca alguns.
  const all=[...DATA.professores];
  const visible=new Set(matched.map(p=>p.id));
  const statusOrder=['Ativos','Colaboradores','Aposentados','Ex-professores','In Memoriam'];
  all.sort((a,b)=>{
    const ia=statusOrder.indexOf(a.status),ib=statusOrder.indexOf(b.status);
    return (ia<0?99:ia)-(ib<0?99:ib)||a.nome.localeCompare(b.nome,'pt-BR');
  });
  const seats=seatsHemicircle(all.length);
  const cx=460,cy=452;
  // Linhas-guia pouco contrastantes ajudam a reconhecer o hemiciclo.
  for(const radius of [112,152,192,232,272,312,352]){
    svg.appendChild(makeSVG('path',{
      d:`M ${cx-radius} ${cy} A ${radius} ${radius} 0 0 1 ${cx+radius} ${cy}`,
      class:'hemicycle-guide',fill:'none'
    }));
  }
  svg.appendChild(makeSVG('path',{d:`M 88 ${cy+4} H 832`,class:'hemicycle-floor',fill:'none'}));
  all.forEach((p,i)=>{
    const pos=seats[i],inFilter=visible.has(p.id),n=countRole(p),selected=p.id===SELECTED;
    const active=inFilter&&n>0;
    const circle=makeSVG('circle',{
      cx:pos.x.toFixed(2),cy:pos.y.toFixed(2),r:selected?11:9.7,
      fill:STATUS_COLORS[p.status]||'#a9afa5',
      opacity:active?'1':inFilter?'.38':'.14',
      stroke:selected?'#152f29':'#f8f4ec',
      'stroke-width':selected?3.2:1.4,
      class:'person-dot hemicycle-seat'+(selected?' selected':'')+(active?' active-seat':' muted-seat'),
      tabindex:inFilter?'0':'-1',role:'button',
      'aria-label':`${p.nome}: ${n} participações no filtro; situação: ${p.status}${inFilter?'':'. Fora dos filtros selecionados'}`
    });
    circle.appendChild(makeSVG('title',{},`${p.nome}\n${n} participações no recorte\n${p.status}`));
    if(inFilter){
      circle.addEventListener('click',()=>{SELECTED=p.id;renderBancas();});
      circle.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();SELECTED=p.id;renderBancas();}});
    }
    svg.appendChild(circle);
  });
  svg.appendChild(makeSVG('text',{
    x:460,y:485,'text-anchor':'middle',class:'hemicycle-caption'
  },`${fmt(all.length)} DOCENTES · CADA ASSENTO É UMA PESSOA`));
  $('mapa-legenda-curta').textContent=`${fmt(matched.length)} DESTACADOS / ${fmt(all.length)} NO MAPA`;
  if(!matched.length){
    svg.appendChild(makeSVG('text',{x:460,y:418,'text-anchor':'middle',class:'hemicycle-empty'},'Nenhuma pessoa corresponde aos filtros.'));
  }
}
function renderSelected(){
  const p=getProfessor(SELECTED);if(!p){$('perfil-nome').textContent='Nenhum professor no filtro';return;}
  const cnt=participationCounts(p),total=countRole(p);
  $('perfil-nome').textContent=p.nome;$('perfil-avatar').textContent=initials(p.nome);
  $('perfil-status').textContent=`Situação no cadastro de 2026: ${p.status}${p.ms&&p.ms!=='Não informado'?' · '+p.ms:''}`;
  $('perfil-total').textContent=fmt(total);$('perfil-bancas').textContent=fmt(p.bancasDistintas);
  bars('perfil-funcoes',cnt,{colors:Object.fromEntries(ROLES.map(k=>[k,'#f5a684']))});
  $('perfil-anos').textContent=`Período do filtro: ${FILTER.ano==='todos'?'2023 a 2025':FILTER.ano}. Total da tabela consolidada: ${p.total} participações.`;
  $('perfil-aviso').textContent=p.vinculosInferidos?`Há ${p.vinculosInferidos} participações com vínculo IC inferido. O status é o do cadastro consultado, não o da data da defesa.`:'O status é o do cadastro consultado, não o da data da defesa. Gênero e raça/cor nominais não são exibidos por falta de confirmação pública.';
}
function initBancas(){
  const s=$('banca-status');for(const status of Object.keys(STATUS_COLORS))s.add(new Option(status,status));
  $('banca-ano').addEventListener('change',e=>{FILTER.ano=e.target.value;renderBancas();});
  $('banca-papel').addEventListener('change',e=>{FILTER.papel=e.target.value;renderBancas();});
  s.addEventListener('change',e=>{FILTER.status=e.target.value;renderBancas();});
  $('banca-busca').addEventListener('input',e=>{FILTER.busca=e.target.value.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();renderBancas();});
  $('banca-legenda').innerHTML=Object.entries(STATUS_COLORS).map(([k,c])=>`<span class="legend-item"><i class="legend-swatch" style="background:${c}"></i>${esc(k)}</span>`).join('');
  const audit=DATA.auditoria;
  if(audit.quantidadeDivergencias){$('banca-auditoria').hidden=false;
    $('banca-auditoria').innerHTML=`<strong>Conferência necessária</strong><p>${fmt(audit.quantidadeDivergencias)} diferenças entre os detalhes gerados a partir do Excel enviado e a tabela consolidada. Nos totais usamos a tabela consolidada. Ao selecionar um ano, contamos os registros do Excel. Nomes para revisão: ${esc(audit.divergencias.map(d=>d.professor).join(', '))}.</p>`;}
  SELECTED=DATA.professores.find(p=>p.total>0)?.id;renderBancas();
}

// 04. PRÊMIOS — prêmios, finalistas e outras distinções não são sinônimos.
function awardsInYear(p){return p.premios.filter(a=>PREMIO_ANO==='todos'||(PREMIO_ANO==='sem-ano'?a.ano===null:String(a.ano)===PREMIO_ANO));}
function countedAwards(p){return awardsInYear(p).filter(a=>a.tipo==='Prêmio/menção informada');}
function trophyLevel(fraction){
  if(TROPHY_REQUEST!==null)cancelAnimationFrame(TROPHY_REQUEST);
  const start=performance.now(),from=TROPHY;
  function frame(now){const t=Math.min((now-start)/650,1),e=1-(1-t)**3;const val=from+(fraction-from)*e;const h=236*val,y=286-h;
    $('trophy-fill').setAttribute('height',h.toFixed(2));$('trophy-fill').setAttribute('y',y.toFixed(2));
    $('trophy-wave').setAttribute('d',`M96 ${y+2} C134 ${y-8} 180 ${y+10} 224 ${y+1}`);
    if(t<1)TROPHY_REQUEST=requestAnimationFrame(frame);else{TROPHY=val;TROPHY_REQUEST=null;}}
  if(window.matchMedia('(prefers-reduced-motion:reduce)').matches){TROPHY=fraction;const h=236*fraction;$('trophy-fill').setAttribute('height',h);$('trophy-fill').setAttribute('y',286-h);return;}
  TROPHY_REQUEST=requestAnimationFrame(frame);
}
function renderAwards(){
  const p=getProfessor(PREMIO_PERSON);if(!p)return;
  const items=awardsInYear(p),actual=countedAwards(p),max=Math.max(1,...DATA.professores.map(p=>p.premiosNumero));
  $('premio-num').textContent=fmt(actual.length);$('premio-escala').textContent=`Escala ilustrativa. Maior contagem na tabela: ${max} registros para uma pessoa.`;
  trophyLevel(Math.min(1,actual.length/max));
  $('premio-meta').innerHTML=`<span class="meta-pill">Cadastro: ${esc(p.status)}</span><span class="meta-pill">${actual.length} prêmios / menções</span><span class="meta-pill">${items.length-actual.length} outras distinções</span>`;
  $('lista-premios').innerHTML=items.length?items.map(a=>`<div class="award-item"><time>${a.ano??'S/A'}</time><div><h4>${esc(a.titulo)}</h4><span>${esc(a.tipo)}</span></div></div>`).join(''):'<p class="award-empty">Nenhum registro de reconhecimento encontrado para o filtro.</p>';
  const agg=DATA.distribuicoes.premios;
  bars('premio-dist-genero',agg.genero);bars('premio-dist-raca',agg.raca);bars('premio-dist-status',agg.status,{colors:STATUS_COLORS});
}
function initAwards(){
  const people=[...DATA.professores].sort((a,b)=>a.nome.localeCompare(b.nome,'pt-BR'));
  const s=$('premio-pessoa');people.forEach(p=>s.add(new Option(p.nome,p.id)));
  const yearSel=$('premio-ano');const years=[...new Set(people.flatMap(p=>p.premios.map(x=>x.ano).filter(Boolean)))].sort((a,b)=>b-a);
  years.forEach(y=>yearSel.add(new Option(String(y),String(y))));yearSel.add(new Option('Ano não informado','sem-ano'));
  PREMIO_PERSON=[...people].sort((a,b)=>b.premiosNumero-a.premiosNumero)[0]?.id;
  s.value=PREMIO_PERSON;
  s.addEventListener('change',e=>{PREMIO_PERSON=e.target.value;renderAwards();});
  yearSel.addEventListener('change',e=>{PREMIO_ANO=e.target.value;renderAwards();});
  renderAwards();
}

// 05. EXPERIMENTO — não é um mecanismo de avaliação de docentes.
function selectedGameFields(){return [
  ['Orientador(a)','orientador','Orientação'],['Coorientador(a)','coorientador','Coorientação'],
  ...[1,2,3,4,5].map(i=>[`Titular ${i}`,`titular-${i}`,'Titular']),
  ...[1,2,3].map(i=>[`Suplente ${i}`,`suplente-${i}`,'Suplente'])
];}
function renderGame(){
  const ids=[],fields=selectedGameFields();let total=0;
  $('cadeiras-selecionadas').innerHTML=fields.map(([label,id,role])=>{
    const p=getProfessor($('cadeira-'+id).value),n=p?.funcoes[role]||0;if(p){ids.push(p.id);total+=n;}
    return `<div class="chair-pill"><span class="func">${esc(label.toUpperCase())}</span><strong class="${p?'':'empty'}">${p?esc(p.nome):'Cadeira vazia'}</strong>${p?`<span class="count">${fmt(n)} participações na função</span>`:''}</div>`;
  }).join('');
  const unique=new Set(ids).size;
  $('texto-banca').textContent=ids.length?`${ids.length} ${ids.length===1?'cadeira preenchida':'cadeiras preenchidas'} por ${unique} ${unique===1?'pessoa':'pessoas'}. ${ids.length===1?'Essa escolha soma':'Juntas, essas escolhas somam'} ${fmt(total)} participações históricas nas funções correspondentes. ${unique<ids.length?'Há uma pessoa selecionada em mais de uma função. ':''}É um exercício exploratório, não uma indicação para uma banca real.`:'Escolha uma pessoa para começar a experiência.';
}
function initGame(){
  for(let i=1;i<=5;i++)$('titulares-inputs').insertAdjacentHTML('beforeend',`<label>TITULAR ${i}<select id="cadeira-titular-${i}"></select></label>`);
  for(let i=1;i<=3;i++)$('suplentes-inputs').insertAdjacentHTML('beforeend',`<label>SUPLENTE ${i}<select id="cadeira-suplente-${i}"></select></label>`);
  const people=[...DATA.professores].sort((a,b)=>a.nome.localeCompare(b.nome,'pt-BR'));
  for(const [_,id] of selectedGameFields()){
    const select=$('cadeira-'+id);select.add(new Option('Selecione...',''));
    people.forEach(p=>select.add(new Option(p.nome,p.id)));
    select.addEventListener('change',renderGame);
  }
  $('limpar-banca').addEventListener('click',()=>{selectedGameFields().forEach(([_,id])=>$('cadeira-'+id).value='');renderGame();notify('Seleção limpa.');});
  renderGame();
}
function readingProgress(){const max=document.documentElement.scrollHeight-innerHeight;const p=max>0?scrollY/max*100:0;$('reading-line').style.width=p+'%';}
async function boot(){
  heroDots();
  try{
    const r=await fetch('/api/dados',{cache:'no-store'});
    const json=await r.json();
    if(!r.ok)throw new Error(json.erro||'Falha na leitura dos dados');
    DATA=json;
    initDocentes();initBancas();initAwards();initGame();
    document.addEventListener('scroll',readingProgress,{passive:true});readingProgress();closeLoading();
  }catch(error){console.error(error);showError(`Não foi possível iniciar o site: ${error.message}. Confira o terminal do VS Code e as planilhas na pasta dados.`);}
}
document.addEventListener('DOMContentLoaded',boot);
