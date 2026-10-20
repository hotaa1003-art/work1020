// 국방 예산안 대시보드 스크립트

// 기본 샘플 데이터 (업로드된 이미지의 데이터와 100% 일치)
const SAMPLE_DATA = [
  {
    id: 1,
    category: "① 자주국방 & 전작권 전환",
    name: "전작권 전환 및 자주국방 핵심전력 확보",
    unit: "억원",
    budget2026: 90222,
    budget2027: 121744,
    diff: 31522,
    rate: "34.9%",
    rateVal: 34.9,
    isNew: false,
    notes: "전작권 전환 및 자주국방 핵심전력 확보"
  },
  {
    id: 2,
    category: "② 효율적 군 인력구조 전환",
    name: "상비예비군 운영 예산",
    unit: "억원",
    budget2026: 87,
    budget2027: 256,
    diff: 169,
    rate: "194.3%",
    rateVal: 194.3,
    isNew: false,
    notes: "평시 추가훈련을 통한 전투력 및 전문성 강화"
  },
  {
    id: 3,
    category: "② 효율적 군 인력구조 전환",
    name: "예비군 훈련비 총예산",
    unit: "억원",
    budget2026: 688,
    budget2027: 903,
    diff: 215,
    rate: "31.3%",
    rateVal: 31.3,
    isNew: false,
    notes: "'30년까지 최저시급 수준 단계적 인상"
  },
  {
    id: 4,
    category: "③ AI·유무인 복합체계(MUM-T)",
    name: "AI·유무인 복합체계 총예산",
    unit: "억원",
    budget2026: 4968,
    budget2027: 7727,
    diff: 2759,
    rate: "55.5%",
    rateVal: 55.5,
    isNew: false,
    notes: "첨단 유·무인 복합체계 조기 전력화 및 핵심기술 개발"
  },
  {
    id: 5,
    category: "③ AI·유무인 복합체계(MUM-T)",
    name: "피지컬 AI 실증사업",
    unit: "억원",
    budget2026: 0,
    budget2027: 430,
    diff: 430,
    rate: "신규",
    rateVal: null,
    isNew: true,
    notes: "[신규 430억원] 수송·주둔지 경계 로봇 등 피지컬 AI 실증"
  },
  {
    id: 6,
    category: "④ 드론 / 대드론 역량강화",
    name: "드론·대드론 역량강화 예산",
    unit: "억원",
    budget2026: 6788,
    budget2027: 8061,
    diff: 1273,
    rate: "18.8%",
    rateVal: 18.8,
    isNew: false,
    notes: "공격형 드론 확충 및 다층 대드론 방호체계 구축"
  }
];

// 전역 상태
let currentData = JSON.parse(JSON.stringify(SAMPLE_DATA));
let filteredData = [];
let sortKey = 'id';
let sortDirection = 'asc';
let activeCategoryFilter = 'ALL';
let charts = {};

// 숫자 포맷 함수
function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return '-';
  return Number(num).toLocaleString('ko-KR');
}

// 초기화
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  loadData(currentData, '기본 제공 국방예산 샘플 데이터 (6건)');
});

// 이벤트 리스너 등록
function setupEventListeners() {
  // 탭 전환
  const tabBtns = document.querySelectorAll('.tab-btn');
  tabBtns.forEach(btn => {
    btn.addEventListener('click', (e) => {
      tabBtns.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.dashboard-section').forEach(s => s.classList.remove('active'));

      const target = btn.dataset.tab;
      btn.classList.add('active');
      const section = document.getElementById(target);
      if (section) section.classList.add('active');
    });
  });

  // 파일 업로드 (Drag & Drop 및 인풋)
  const dropZone = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  dropZone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleFileSelect);

  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });

  ['dragleave', 'dragend'].forEach(type => {
    dropZone.addEventListener(type, () => dropZone.classList.remove('dragover'));
  });

  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      processFile(e.dataTransfer.files[0]);
    }
  });

  // 샘플 데이터 다시 로드
  document.getElementById('btnReloadSample')?.addEventListener('click', () => {
    currentData = JSON.parse(JSON.stringify(SAMPLE_DATA));
    loadData(currentData, '기본 샘플 데이터 재로드 완료');
  });

  // 서식 다운로드 (CSV / Excel)
  document.getElementById('btnDownloadSample')?.addEventListener('click', downloadSampleCSV);
  document.getElementById('btnExportExcel')?.addEventListener('click', exportToExcel);
  document.getElementById('btnExportCSV')?.addEventListener('click', exportToCSV);

  // 테이블 검색 및 필터
  const searchInput = document.getElementById('searchInput');
  const catFilter = document.getElementById('categoryFilter');
  const onlyNewToggle = document.getElementById('onlyNewToggle');

  if (searchInput) searchInput.addEventListener('input', applyFilters);
  if (catFilter) catFilter.addEventListener('change', applyFilters);
  if (onlyNewToggle) onlyNewToggle.addEventListener('change', applyFilters);

  // 테이블 헤더 정렬
  document.querySelectorAll('table.budget-table th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.sort;
      if (sortKey === key) {
        sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
      } else {
        sortKey = key;
        sortDirection = 'asc';
      }
      applyFilters();
    });
  });
}

// 파일 선택 처리
function handleFileSelect(e) {
  if (e.target.files.length > 0) {
    processFile(e.target.files[0]);
  }
}

// 파일 파싱 (XLSX, XLS, CSV 지원)
function processFile(file) {
  const fileName = file.name;
  const ext = fileName.split('.').pop().toLowerCase();

  const reader = new FileReader();

  if (ext === 'xlsx' || ext === 'xls') {
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        if (typeof XLSX === 'undefined') {
          alert('Excel 파싱 라이브러리를 로드할 수 없습니다.');
          return;
        }
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rawJson = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        parseRawTableData(rawJson, fileName);
      } catch (err) {
        console.error(err);
        alert('엑셀 파일 파싱 중 오류가 발생했습니다: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  } else if (ext === 'csv') {
    // CSV 파싱
    reader.onload = (e) => {
      try {
        const text = e.target.result;
        parseCSVText(text, fileName);
      } catch (err) {
        console.error(err);
        alert('CSV 파일 파싱 중 오류가 발생했습니다: ' + err.message);
      }
    };
    // EUC-KR / UTF-8 시도
    reader.readAsText(file, 'utf-8');
  } else {
    alert('지원하지 않는 파일 형식입니다. .xlsx, .xls, .csv 파일을 선택해주세요.');
  }
}

// CSV 텍스트 파서
function parseCSVText(csvText, fileName) {
  const lines = csvText.split(/\r\n|\n/).filter(line => line.trim() !== '');
  const rows = lines.map(line => {
    const row = [];
    let inQuotes = false;
    let field = '';
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === ',' && !inQuotes) {
        row.push(field.trim());
        field = '';
      } else {
        field += char;
      }
    }
    row.push(field.trim());
    return row;
  });
  parseRawTableData(rows, fileName);
}

// 원시 2차원 배열 데이터를 모델 규격으로 정규화
function parseRawTableData(rows, fileName) {
  if (!rows || rows.length < 2) {
    alert('데이터가 충분하지 않거나 헤더가 누락되었습니다.');
    return;
  }

  // 헤더 매핑 찾기
  let headerIndex = -1;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const rowStr = rows[i].join(' ');
    if (rowStr.includes('분야') || rowStr.includes('사업') || rowStr.includes('2026') || rowStr.includes('2027')) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) {
    headerIndex = 0;
  }

  const headerRow = rows[headerIndex].map(h => String(h || '').trim());
  const colMap = {
    id: -1,
    category: -1,
    name: -1,
    unit: -1,
    budget2026: -1,
    budget2027: -1,
    diff: -1,
    rate: -1,
    notes: -1
  };

  headerRow.forEach((h, idx) => {
    const colName = h.toLowerCase().replace(/\s+/g, '');
    if (/연번|순번|번호|no|id/.test(colName) && colMap.id === -1) colMap.id = idx;
    else if (/분야|중점영역|중점분야|분야\(중점영역\)|category/.test(colName) && colMap.category === -1) colMap.category = idx;
    else if (/세부사업|주요지표|사업명|사업|project|name/.test(colName) && colMap.name === -1) colMap.name = idx;
    else if (/구분|단위|구분\/단위|unit/.test(colName) && colMap.unit === -1) colMap.unit = idx;
    else if (/2026|전년|기존|2026년\(a\)/.test(colName) && colMap.budget2026 === -1) colMap.budget2026 = idx;
    else if (/2027|정부안|당해|2027년정부안\(b\)/.test(colName) && colMap.budget2027 === -1) colMap.budget2027 = idx;
    else if (/증감|b-a|차이|diff/.test(colName) && colMap.diff === -1) colMap.diff = idx;
    else if (/증가율|증감율|변화율|%|rate/.test(colName) && colMap.rate === -1) colMap.rate = idx;
    else if (/핵심|내용|비고|설명|주요내용|notes/.test(colName) && colMap.notes === -1) colMap.notes = idx;
  });

  const parsed = [];
  for (let r = headerIndex + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0 || row.every(c => c === '' || c === undefined || c === null)) continue;

    const rawId = colMap.id !== -1 ? row[colMap.id] : r - headerIndex;
    const category = colMap.category !== -1 ? String(row[colMap.category] || '기타').trim() : '기타 분야';
    const name = colMap.name !== -1 ? String(row[colMap.name] || '').trim() : `사업 ${r}`;
    const unit = colMap.unit !== -1 ? String(row[colMap.unit] || '억원').trim() : '억원';
    
    // 금액 정제
    const raw2026 = colMap.budget2026 !== -1 ? row[colMap.budget2026] : 0;
    const raw2027 = colMap.budget2027 !== -1 ? row[colMap.budget2027] : 0;
    
    const budget2026 = cleanNumber(raw2026);
    const budget2027 = cleanNumber(raw2027);

    // 증감액 계산
    let diff = colMap.diff !== -1 && row[colMap.diff] !== undefined && row[colMap.diff] !== '' 
      ? cleanNumber(row[colMap.diff]) 
      : (budget2027 - budget2026);

    // 증가율 및 신규 여부 파악
    const rawRate = colMap.rate !== -1 ? String(row[colMap.rate] || '').trim() : '';
    const notes = colMap.notes !== -1 ? String(row[colMap.notes] || '').trim() : '';

    let isNew = false;
    let rateStr = '';
    let rateVal = null;

    if (rawRate.includes('신규') || budget2026 === 0 || notes.includes('[신규]') || notes.includes('신규') || name.includes('신규')) {
      isNew = true;
      rateStr = '신규';
    } else {
      if (rawRate && !isNaN(cleanNumber(rawRate))) {
        rateVal = cleanNumber(rawRate);
        rateStr = rateVal.toFixed(1) + '%';
      } else if (budget2026 > 0) {
        rateVal = Number(((diff / budget2026) * 100).toFixed(1));
        rateStr = (rateVal > 0 ? '+' : '') + rateVal.toFixed(1) + '%';
      } else {
        rateStr = '-';
      }
    }

    parsed.push({
      id: rawId || (parsed.length + 1),
      category: category || '기타',
      name: name || '이름 없음',
      unit: unit || '억원',
      budget2026,
      budget2027,
      diff,
      rate: rateStr,
      rateVal,
      isNew,
      notes
    });
  }

  if (parsed.length === 0) {
    alert('유효한 데이터 행을 찾을 수 없습니다. 서식을 확인해 주세요.');
    return;
  }

  currentData = parsed;
  loadData(currentData, `"${fileName}" 파일 로드 완료 (${parsed.length}건)`);
}

function cleanNumber(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === 'number') return val;
  const cleaned = String(val).replace(/,/g, '').replace(/%/g, '').replace(/억원/g, '').trim();
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

// 데이터 로드 및 전체 렌더링
function loadData(data, statusMsg) {
  // 상태 바 업데이트
  const statusBar = document.getElementById('fileStatusBar');
  const statusText = document.getElementById('fileStatusText');
  if (statusBar && statusText) {
    statusBar.style.display = 'flex';
    statusText.innerHTML = `<strong>성공:</strong> ${statusMsg} | 총 ${data.length}개 사업 항목 로드됨`;
  }

  // 분야 필터 옵션 업데이트
  updateCategoryFilterOptions(data);

  // 필터 적용 및 화면 갱신
  applyFilters();
}

// 분야 셀렉트박스 옵션 채우기
function updateCategoryFilterOptions(data) {
  const select = document.getElementById('categoryFilter');
  if (!select) return;

  const currentVal = select.value;
  const categories = [...new Set(data.map(d => d.category))];

  select.innerHTML = '<option value="ALL">전체 분야 보기</option>';
  categories.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.textContent = cat;
    select.appendChild(opt);
  });

  if (categories.includes(currentVal)) {
    select.value = currentVal;
  }
}

// 필터 및 정렬 적용
function applyFilters() {
  const searchVal = (document.getElementById('searchInput')?.value || '').toLowerCase().trim();
  const catVal = document.getElementById('categoryFilter')?.value || 'ALL';
  const onlyNew = document.getElementById('onlyNewToggle')?.checked || false;

  filteredData = currentData.filter(item => {
    if (catVal !== 'ALL' && item.category !== catVal) return false;
    if (onlyNew && !item.isNew) return false;
    if (searchVal) {
      const matchName = item.name.toLowerCase().includes(searchVal);
      const matchCat = item.category.toLowerCase().includes(searchVal);
      const matchNotes = item.notes.toLowerCase().includes(searchVal);
      if (!matchName && !matchCat && !matchNotes) return false;
    }
    return true;
  });

  // 정렬 적용
  filteredData.sort((a, b) => {
    let valA = a[sortKey];
    let valB = b[sortKey];

    if (sortKey === 'rate') {
      valA = a.rateVal !== null ? a.rateVal : (a.isNew ? 999999 : -999999);
      valB = b.rateVal !== null ? b.rateVal : (b.isNew ? 999999 : -999999);
    }

    if (typeof valA === 'string') {
      return sortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
    } else {
      return sortDirection === 'asc' ? (valA - valB) : (valB - valA);
    }
  });

  // KPI 요약 카드 렌더링
  renderKPICards(currentData);

  // 신규 사업 분석 섹션 렌더링
  renderNewProjectsSection(currentData);

  // 분야별 카드 및 통계 렌더링
  renderCategorySection(currentData);

  // 차트 렌더링
  renderCharts(currentData);

  // 데이터 테이블 렌더링
  renderTable(filteredData);
}

// KPI 요약 지표 카드 계산 및 출력
function renderKPICards(data) {
  const sum2026 = data.reduce((acc, d) => acc + d.budget2026, 0);
  const sum2027 = data.reduce((acc, d) => acc + d.budget2027, 0);
  const totalDiff = sum2027 - sum2026;
  const totalRate = sum2026 > 0 ? ((totalDiff / sum2026) * 100).toFixed(1) : 0;

  const newProjects = data.filter(d => d.isNew);
  const newBudget = newProjects.reduce((acc, d) => acc + d.budget2027, 0);
  const newPct = sum2027 > 0 ? ((newBudget / sum2027) * 100).toFixed(1) : 0;

  document.getElementById('kpiBudget2026').textContent = formatNumber(sum2026);
  document.getElementById('kpiBudget2027').textContent = formatNumber(sum2027);
  
  const diffElem = document.getElementById('kpiTotalDiff');
  diffElem.textContent = (totalDiff >= 0 ? '+' : '') + formatNumber(totalDiff);
  
  const rateElem = document.getElementById('kpiTotalRate');
  rateElem.textContent = (totalRate >= 0 ? '+' : '') + totalRate + '%';

  document.getElementById('kpiNewCount').textContent = `${newProjects.length}건`;
  document.getElementById('kpiNewBudget').textContent = `${formatNumber(newBudget)} 억원 (${newPct}%)`;

  // 탭 배지 숫자
  document.getElementById('badgeTotalCount').textContent = `${data.length}개 사업`;
  document.getElementById('badgeNewCount').textContent = `${newProjects.length}건 신규`;
}

// 신규 사업 집중 분석 섹션 렌더링
function renderNewProjectsSection(data) {
  const newItems = data.filter(d => d.isNew);
  const container = document.getElementById('newProjectsGrid');
  const summaryText = document.getElementById('spotlightSummaryText');

  const sumNewBudget = newItems.reduce((acc, d) => acc + d.budget2027, 0);
  if (summaryText) {
    summaryText.innerHTML = `총 <strong>${newItems.length}개 신규 사업</strong>에 <strong>${formatNumber(sumNewBudget)} 억원</strong>이 2027년 정부안으로 신규 투입됩니다.`;
  }

  if (!container) return;
  container.innerHTML = '';

  if (newItems.length === 0) {
    container.innerHTML = `<div style="grid-column: 1/-1; padding: 40px; text-align: center; color: #94a3b8; background: #fff; border-radius: 12px; border: 1px dashed #cbd5e1;">현재 감지된 신규 사업이 없습니다. (2026년 예산이 0원이거나 '신규'로 표기된 사업)</div>`;
    return;
  }

  newItems.forEach(item => {
    const card = document.createElement('div');
    card.className = 'new-project-card';
    card.innerHTML = `
      <span class="new-badge">⭐ 2027 신규 실증·착수 사업</span>
      <h4 class="new-card-title">${item.name}</h4>
      <div class="new-card-cat">${item.category}</div>
      <div class="new-card-budget">
        <span style="font-size: 0.85rem; color: #78350f; font-weight: 600;">2027년 신규 편성액:</span>
        <span class="new-card-budget-val">${formatNumber(item.budget2027)}</span>
        <span style="font-size: 0.9rem; color: #78350f; font-weight: 600;">${item.unit}</span>
      </div>
      <div class="new-card-notes">
        <strong>주요 추진 계획:</strong><br/>
        ${item.notes || '상세 추진 계획이 등록되지 않았습니다.'}
      </div>
    `;
    container.appendChild(card);
  });
}

// 분야별 분석 카드 그리드
function renderCategorySection(data) {
  const container = document.getElementById('categoryCardsGrid');
  if (!container) return;
  container.innerHTML = '';

  // 분야별 그룹핑
  const catMap = {};
  data.forEach(d => {
    if (!catMap[d.category]) {
      catMap[d.category] = {
        name: d.category,
        count: 0,
        b2026: 0,
        b2027: 0,
        newCount: 0,
        newBudget: 0
      };
    }
    catMap[d.category].count++;
    catMap[d.category].b2026 += d.budget2026;
    catMap[d.category].b2027 += d.budget2027;
    if (d.isNew) {
      catMap[d.category].newCount++;
      catMap[d.category].newBudget += d.budget2027;
    }
  });

  Object.values(catMap).forEach(cat => {
    const diff = cat.b2027 - cat.b2026;
    const rate = cat.b2026 > 0 ? ((diff / cat.b2026) * 100).toFixed(1) : '신규';
    const rateClass = rate > 0 ? 'pos' : (rate === '신규' ? 'new' : 'neg');

    const card = document.createElement('div');
    card.className = 'cat-card' + (activeCategoryFilter === cat.name ? ' selected' : '');
    card.onclick = () => {
      const select = document.getElementById('categoryFilter');
      if (select) {
        select.value = (select.value === cat.name) ? 'ALL' : cat.name;
        applyFilters();
        // 테이블 탭으로 포커스
        const tabTable = document.querySelector('.tab-btn[data-tab="tabTable"]');
        if (tabTable) tabTable.click();
      }
    };

    card.innerHTML = `
      <div class="cat-header">
        <div class="cat-name">${cat.name}</div>
        <span class="cat-count">${cat.count}개 사업</span>
      </div>
      <div class="cat-metrics">
        <div>
          <div class="cat-metric-label">2026년</div>
          <div class="cat-metric-val">${formatNumber(cat.b2026)} 억원</div>
        </div>
        <div>
          <div class="cat-metric-label">2027년 정부안</div>
          <div class="cat-metric-val" style="color: var(--primary);">${formatNumber(cat.b2027)} 억원</div>
        </div>
        <div>
          <div class="cat-metric-label">증감액 (B-A)</div>
          <div class="cat-metric-val" style="color: ${diff >= 0 ? '#10b981' : '#ef4444'}">${(diff >= 0 ? '+' : '') + formatNumber(diff)} 억원</div>
        </div>
        <div>
          <div class="cat-metric-label">증가율</div>
          <div class="cat-metric-val"><span class="rate-badge ${rateClass}">${rate !== '신규' ? (rate + '%') : '신규'}</span></div>
        </div>
      </div>
      ${cat.newCount > 0 ? `
        <div style="margin-top: 10px; padding: 6px 10px; background: #fffbeb; border-radius: 6px; font-size: 0.78rem; color: #92400e; font-weight: 600;">
          ⚡ 신규사업 ${cat.newCount}건 (${formatNumber(cat.newBudget)} 억원) 포함
        </div>
      ` : ''}
    `;
    container.appendChild(card);
  });
}

// Chart.js 시각화 렌더링
function renderCharts(data) {
  if (typeof Chart === 'undefined') {
    console.warn('Chart.js가 로드되지 않았습니다.');
    return;
  }

  // 분야별 통계
  const catMap = {};
  data.forEach(d => {
    if (!catMap[d.category]) {
      catMap[d.category] = { b2026: 0, b2027: 0 };
    }
    catMap[d.category].b2026 += d.budget2026;
    catMap[d.category].b2027 += d.budget2027;
  });

  const catLabels = Object.keys(catMap);
  const data2026 = catLabels.map(c => catMap[c].b2026);
  const data2027 = catLabels.map(c => catMap[c].b2027);

  // 차트 1: 분야별 비교 막대 차트
  const ctxCompare = document.getElementById('chartCategoryCompare')?.getContext('2d');
  if (ctxCompare) {
    if (charts.compare) charts.compare.destroy();
    charts.compare = new Chart(ctxCompare, {
      type: 'bar',
      data: {
        labels: catLabels.map(l => l.length > 15 ? l.substring(0, 15) + '...' : l),
        datasets: [
          {
            label: '2026년 (A)',
            data: data2026,
            backgroundColor: '#94a3b8',
            borderRadius: 6
          },
          {
            label: '2027년 정부안 (B)',
            data: data2027,
            backgroundColor: '#1e3a8a',
            borderRadius: 6
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'top' },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatNumber(ctx.raw)} 억원`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              callback: (val) => formatNumber(val)
            }
          }
        }
      }
    });
  }

  // 차트 2: 분야별 점유율 도넛 차트
  const ctxShare = document.getElementById('chartCategoryShare')?.getContext('2d');
  if (ctxShare) {
    if (charts.share) charts.share.destroy();
    charts.share = new Chart(ctxShare, {
      type: 'doughnut',
      data: {
        labels: catLabels,
        datasets: [{
          data: data2027,
          backgroundColor: [
            '#1e3a8a',
            '#3b82f6',
            '#f59e0b',
            '#10b981',
            '#8b5cf6',
            '#ec4899',
            '#64748b'
          ]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right' },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const total = data2027.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : 0;
                return `${ctx.label}: ${formatNumber(ctx.raw)} 억원 (${pct}%)`;
              }
            }
          }
        }
      }
    });
  }

  // 차트 3: 신규 사업 vs 기존 사업 비중
  const ctxNewShare = document.getElementById('chartNewVsExisting')?.getContext('2d');
  if (ctxNewShare) {
    if (charts.newShare) charts.newShare.destroy();
    const newSum = data.filter(d => d.isNew).reduce((acc, d) => acc + d.budget2027, 0);
    const existSum = data.filter(d => !d.isNew).reduce((acc, d) => acc + d.budget2027, 0);

    charts.newShare = new Chart(ctxNewShare, {
      type: 'pie',
      data: {
        labels: ['기존 사업 증액/유지', '2027년 신규 사업'],
        datasets: [{
          data: [existSum, newSum],
          backgroundColor: ['#1e40af', '#f59e0b']
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const total = newSum + existSum;
                const pct = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : 0;
                return `${ctx.label}: ${formatNumber(ctx.raw)} 억원 (${pct}%)`;
              }
            }
          }
        }
      }
    });
  }

  // 차트 4: 증가율 상위 사업 랭킹
  const ctxRanking = document.getElementById('chartRateRanking')?.getContext('2d');
  if (ctxRanking) {
    if (charts.ranking) charts.ranking.destroy();
    
    // 증가율이 유효한 항목 순위 (신규 제외 상위 5개)
    const ranked = [...data]
      .filter(d => d.rateVal !== null && !isNaN(d.rateVal))
      .sort((a, b) => b.rateVal - a.rateVal)
      .slice(0, 6);

    charts.ranking = new Chart(ctxRanking, {
      type: 'bar',
      data: {
        labels: ranked.map(d => d.name.length > 14 ? d.name.substring(0, 14) + '...' : d.name),
        datasets: [{
          label: '증가율 (%)',
          data: ranked.map(d => d.rateVal),
          backgroundColor: ranked.map((d, i) => i === 0 ? '#10b981' : '#3b82f6'),
          borderRadius: 6
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `증가율: +${ctx.raw}%`
            }
          }
        },
        scales: {
          x: {
            ticks: {
              callback: (val) => `+${val}%`
            }
          }
        }
      }
    });
  }
}

// 상세 데이터 테이블 렌더링
function renderTable(data) {
  const tbody = document.getElementById('budgetTableBody');
  const tableCount = document.getElementById('tableRowCount');
  const footerSummary = document.getElementById('tableFooterSummary');

  if (!tbody) return;
  tbody.innerHTML = '';

  if (tableCount) tableCount.textContent = `${data.length} 건`;

  if (data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align: center; padding: 40px; color: #94a3b8;">일치하는 사업 내역이 없습니다.</td></tr>`;
    if (footerSummary) footerSummary.innerHTML = '';
    return;
  }

  let total2026 = 0;
  let total2027 = 0;
  let totalDiff = 0;

  data.forEach((item, index) => {
    total2026 += item.budget2026;
    total2027 += item.budget2027;
    totalDiff += item.diff;

    const tr = document.createElement('tr');
    if (item.isNew) tr.classList.add('row-new');

    const rateBadgeClass = item.isNew ? 'new' : (item.diff >= 0 ? 'pos' : 'neg');
    const rateText = item.isNew ? '신규' : item.rate;

    tr.innerHTML = `
      <td style="text-align: center; font-weight: 700;">${item.id}</td>
      <td><span class="badge-cat-tag">${item.category}</span></td>
      <td>
        <div style="font-weight: 700; color: #0f172a;">${item.name}</div>
      </td>
      <td style="text-align: center; color: #64748b;">${item.unit}</td>
      <td class="cell-num">${formatNumber(item.budget2026)}</td>
      <td class="cell-num" style="color: var(--primary); font-weight: 800;">${formatNumber(item.budget2027)}</td>
      <td class="cell-num" style="color: ${item.diff >= 0 ? '#10b981' : '#ef4444'};">
        ${(item.diff >= 0 ? '+' : '') + formatNumber(item.diff)}
      </td>
      <td style="text-align: center;">
        <span class="rate-badge ${rateBadgeClass}">${rateText}</span>
      </td>
      <td style="font-size: 0.82rem; color: #475569;">
        ${item.isNew ? '<span class="badge-new-pill" style="margin-right: 4px;">신규</span>' : ''}
        ${item.notes || '-'}
      </td>
    `;
    tbody.appendChild(tr);
  });

  // 테이블 하단 합계 요약
  if (footerSummary) {
    const overallRate = total2026 > 0 ? ((totalDiff / total2026) * 100).toFixed(1) + '%' : '-';
    footerSummary.innerHTML = `
      <div>조회된 <strong>${data.length}</strong>개 사업 합계</div>
      <div style="display: flex; gap: 24px; font-weight: 700;">
        <span>2026년 합계: <strong>${formatNumber(total2026)} 억원</strong></span>
        <span>2027년 정부안 합계: <strong style="color: var(--primary);">${formatNumber(total2027)} 억원</strong></span>
        <span>순증감: <strong style="color: #10b981;">+${formatNumber(totalDiff)} 억원 (${overallRate})</strong></span>
      </div>
    `;
  }
}

// 샘플 CSV 파일 다운로드 생성
function downloadSampleCSV() {
  const headers = ['연번', '분야 (중점영역)', '세부 사업 / 주요 지표', '구분/단위', '2026년 (A)', '2027년 정부안 (B)', '증감 (B-A)', '증가율 (%)', '핵심 내용'];
  const rows = SAMPLE_DATA.map(d => [
    d.id,
    `"${d.category}"`,
    `"${d.name}"`,
    d.unit,
    d.budget2026,
    d.budget2027,
    d.diff,
    `"${d.rate}"`,
    `"${d.notes}"`
  ]);

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '국방예산안_서식_샘플.csv';
  a.click();
  URL.revokeObjectURL(url);
}

// 엑셀 내보내기 (.xlsx)
function exportToExcel() {
  if (typeof XLSX === 'undefined') {
    alert('Excel 라이브러리를 찾을 수 없습니다.');
    return;
  }

  const exportRows = filteredData.map(d => ({
    '연번': d.id,
    '분야 (중점영역)': d.category,
    '세부 사업 / 주요 지표': d.name,
    '구분/단위': d.unit,
    '2026년 (A)': d.budget2026,
    '2027년 정부안 (B)': d.budget2027,
    '증감 (B-A)': d.diff,
    '증가율 (%)': d.rate,
    '핵심 내용': d.notes
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '예산안_대시보드');
  XLSX.writeFile(workbook, '국방예산안_데이터_' + new Date().toISOString().slice(0, 10) + '.xlsx');
}

// 현재 테이블 CSV 내보내기
function exportToCSV() {
  const headers = ['연번', '분야 (중점영역)', '세부 사업 / 주요 지표', '구분/단위', '2026년 (A)', '2027년 정부안 (B)', '증감 (B-A)', '증가율 (%)', '핵심 내용'];
  const rows = filteredData.map(d => [
    d.id,
    `"${d.category}"`,
    `"${d.name}"`,
    d.unit,
    d.budget2026,
    d.budget2027,
    d.diff,
    `"${d.rate}"`,
    `"${d.notes}"`
  ]);

  const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '국방예산안_데이터_' + new Date().toISOString().slice(0, 10) + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}
